
from datetime import timedelta
import os
import json
import fitz
import base64
import re
from pptx import Presentation
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from openai import OpenAI
from comtypes import client as com_client
from pydantic import BaseModel
from typing import Optional, Dict
import uuid
from datetime import datetime
from pathlib import Path
import shutil
from langfuse import Langfuse
from dotenv import load_dotenv
# FastAPI imports
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import threading
from contextlib import asynccontextmanager

from cleanup import AutoCleanupScheduler
load_dotenv()
langfuse = Langfuse(
    public_key=os.getenv("LANGFUSE_PUBLIC_KEY"),
    secret_key=os.getenv("LANGFUSE_SECRET_KEY"),
    # base_url="https://cloud.langfuse.com"  # or your self-hosted URL
)
class JobManager:
    def __init__(self, storage_file="jobs.json"):
        self._jobs: Dict[str, dict] = {}
        self._lock = threading.Lock()
        self.storage_file = storage_file
        self._load_jobs()
    
    def _load_jobs(self):
        """Load jobs from disk"""
        try:
            if os.path.exists(self.storage_file):
                with open(self.storage_file, 'r') as f:
                    self._jobs = json.load(f)
                print(f"✅ Loaded {len(self._jobs)} jobs from storage")
        except Exception as e:
            print(f"⚠️ Could not load jobs: {e}")
            self._jobs = {}
    
    def _save_jobs(self):
        """Save jobs to disk"""
        try:
            with open(self.storage_file, 'w') as f:
                json.dump(self._jobs, f, indent=2)
        except Exception as e:
            print(f"⚠️ Could not save jobs: {e}")
    
    def create_job(self, job_id: str, job_data: dict):
        with self._lock:
            self._jobs[job_id] = job_data
            self._save_jobs()
    
    def get_job(self, job_id: str) -> Optional[dict]:
        with self._lock:
            return self._jobs.get(job_id)
    
    def update_job(self, job_id: str, updates: dict):
        with self._lock:
            if job_id in self._jobs:
                self._jobs[job_id].update(updates)
                self._save_jobs()
    
    def delete_job(self, job_id: str):
        with self._lock:
            if job_id in self._jobs:
                del self._jobs[job_id]
                self._save_jobs()
    
    def list_jobs(self) -> list:
        with self._lock:
            return list(self._jobs.values())
    
    def get_active_count(self) -> int:
        with self._lock:
            return sum(1 for job in self._jobs.values() if job["status"] == "processing")

# STEP 2: Initialize job manager with persistence
job_manager = JobManager(storage_file="jobs.json")

cleanup_scheduler = AutoCleanupScheduler(
    job_manager=job_manager,
    cleanup_interval_minutes=120,  # Run cleanup every 60 minutes
    job_lifetime_hours=2,         # Delete jobs older than 1 hour
    upload_dir="uploads",
    output_dir="outputs"
)

# Pydantic Models
class SlideOutput(BaseModel):
    technical_notes: str
    

class JobStatus(BaseModel):
    job_id: str
    status: str
    progress: Optional[str] = None
    created_at: str
    completed_at: Optional[str] = None
    error: Optional[str] = None
    output_file: Optional[str] = None

class JobResponse(BaseModel):
    job_id: str
    message: str
    status: str

class AIHandbookMakerStudent:
    def __init__(self, ppt_path, api_key=None):
        self.ppt_path = os.path.abspath(ppt_path)
        self.base_dir = os.path.dirname(self.ppt_path)
        self.prompts_json_path = "prompt_student.json"
        self.classifier_txt_path = "classifier.txt"
        # self.prompts = self.load_json(self.prompts_json_path)
        # self.classifiers = self.load_classifier(self.classifier_txt_path)
        self.pdf_path = os.path.join(self.base_dir, "slides.pdf")
        self.image_dir = os.path.join(self.base_dir, "slide_images")
        self.summary_path = os.path.join(self.base_dir, "summary.txt")
        self.skip_layouts = ['quiz', 'exercise_activity', 'lab_slide']
        
        os.makedirs(self.image_dir, exist_ok=True)
        
        self.doc = Document()
        self.ppt = Presentation(self.ppt_path)
        
        # Load prompts and classifier
        with open(self.prompts_json_path, 'r', encoding='utf-8') as f:
            self.prompts = json.load(f)
        
        with open(self.classifier_txt_path, 'r', encoding='utf-8') as f:
            self.valid_layouts = [l.strip() for l in f.read().split(',')]

        if api_key:
            self.openai_client = OpenAI(api_key=api_key)
        else:
            raise ValueError("❌ OpenAI API key must be provided.")

        self.prompts = {}
        for name in self.valid_layouts:
            prompt_name = f"{name}_student"
            prompt_data = langfuse.get_prompt( name=prompt_name, label="production" )
            if prompt_data and hasattr(prompt_data, "prompt"):
                self.prompts[name] = {"prompt": prompt_data.prompt}
            else:
                print(f"⚠️ Could not fetch prompt for {name}, skipping.")
        
        # Initialize summary file
        self._clear_summary()

    def _clear_summary(self):
        """Clear the summary file"""
        with open(self.summary_path, 'w', encoding='utf-8') as f:
            f.write("")

    def _read_summary(self):
        """Read current summary content"""
        if os.path.exists(self.summary_path):
            with open(self.summary_path, 'r', encoding='utf-8') as f:
                return f.read()
        return ""

    @staticmethod
    def markdown_to_formatted_text(paragraph, text):
        """
        Convert markdown-style formatting to Word formatting
        Handles: **bold**, *italic*
        """
        if not text:
            return
        
        # Pattern to match **bold** and *italic*
        pattern = r'(\*\*.*?\*\*|\*.*?\*)'
        parts = re.split(pattern, text)
        
        for part in parts:
            if not part:
                continue
            if part.startswith('**') and part.endswith('**'):
                # Bold text
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            elif part.startswith('*') and part.endswith('*'):
                # Italic text
                run = paragraph.add_run(part[1:-1])
                run.italic = True
            else:
                # Normal text
                paragraph.add_run(part)

    def _append_summary(self, summary_line):
        """Append a summary line to the file"""
        with open(self.summary_path, 'a', encoding='utf-8') as f:
            f.write(summary_line + "\n")

    def ppt_to_pdf(self):
            """Convert PPTX to PDF using LibreOffice on Ubuntu."""
            print("📄 Converting PPTX to PDF...")

            output_dir = os.path.dirname(self.ppt_path)
            output_pdf = os.path.join(output_dir, "slides.pdf")  # ✅ consistent name

            try:
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", output_dir, self.ppt_path
                ], check=True)

                # Rename generated file to match expected name
                generated_pdf = Path(self.ppt_path).with_suffix(".pdf")
                if os.path.exists(generated_pdf):
                    os.rename(generated_pdf, output_pdf)

                print(f"✅ PDF created successfully at {output_pdf}")
                return output_pdf

            except Exception as e:
                print(f"❌ LibreOffice conversion failed: {e}")
                raise



    def pdf_to_images(self, output_folder="pdf_images"):
        """Convert PDF to images using PyMuPDF"""
        os.makedirs(output_folder, exist_ok=True)
        doc = fitz.open(self.pdf_path)
        image_paths = []

        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap()
            img_path = os.path.join(output_folder, f"page_{i}.png")
            pix.save(img_path)
            image_paths.append(img_path)

        print(f"✅ Converted {len(doc)} pages to images in '{output_folder}'")
        return image_paths

    def extract_text_from_pdf_page(self, page_num):
        """Extract text from a specific PDF page using PyMuPDF"""
        doc = fitz.open(self.pdf_path)
        page = doc[page_num]
        text = page.get_text()
        doc.close()
        return text.strip()

    def extract_heading_from_page(self, page_num):
        """Extract the main heading/title from a PDF page"""
        doc = fitz.open(self.pdf_path)
        page = doc[page_num]
        
        # Get text with formatting info
        blocks = page.get_text("dict")["blocks"]
        
        headings = []
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        size = span["size"]
                        
                        # Consider text as heading if font size > 16 or if it's in upper part of page
                        if text and (size > 16 or line["bbox"][1] < 150):
                            headings.append({
                                "text": text,
                                "size": size,
                                "y_pos": line["bbox"][1]
                            })
        
        doc.close()
        
        # Return the largest/topmost text as heading
        if headings:
            headings.sort(key=lambda x: (-x["size"], x["y_pos"]))
            return headings[0]["text"]
        
        # Fallback: get first line of text
        text = self.extract_text_from_pdf_page(page_num)
        first_line = text.split('\n')[0].strip() if text else ""
        return first_line[:100]

    def identify_topics(self, image_paths):
        """Extract headings from each page and group similar headings into topics"""
        print("🔍 Identifying topics by analyzing page headings...")
        
        total_pages = len(image_paths)
        
        # Extract heading from each page
        page_headings = []
        for i in range(total_pages):
            heading = self.extract_heading_from_page(i)
            page_headings.append({
                "page": i + 1,
                "heading": heading
            })
            print(f"   Page {i + 1}: {heading[:60]}...")
        
        # Prepare prompt with all headings
        headings_text = "\n".join([
            f"Page {p['page']}: {p['heading']}" 
            for p in page_headings
        ])
        prompt = langfuse.get_prompt("identify_topic",label="production")
        prompt=prompt.compile(total_pages=total_pages,headings_text=headings_text)

        
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You analyze slide headings to identify topic groups. Return ONLY valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.5
            )
            
            content = response.choices[0].message.content.strip()
            
            # Extract JSON
            if "```json" in content:
                content = content.split("```json")[1].split("```")[0].strip()
            elif "```" in content:
                content = content.split("```")[1].split("```")[0].strip()
            
            topics_raw = json.loads(content)
            
            # Validate and convert format
            if not topics_raw or not isinstance(topics_raw, list):
                raise ValueError("Invalid topics format")
            
            topics = []
            for t in topics_raw:
                start = t.get("start", 0)
                end = t.get("end", 0)
                
                if start < 1 or end > total_pages or start > end:
                    raise ValueError(f"Invalid page range: {start}-{end}")
                
                count = end - start + 1
                topics.append({
                    "Topic": t["Topic"],
                    "pages": [count, start, end]
                })
            
            # Validate coverage
            if topics[0]["pages"][1] != 1:
                raise ValueError(f"Topics don't start at page 1: {topics[0]}")
            if topics[-1]["pages"][2] != total_pages:
                raise ValueError(f"Topics don't end at page {total_pages}: {topics[-1]}")
            
            # Check gaps
            for i in range(len(topics) - 1):
                if topics[i]["pages"][2] + 1 != topics[i + 1]["pages"][1]:
                    raise ValueError(f"Gap between topics {i} and {i+1}")
            
            print(f"\n✅ Identified {len(topics)} topics:")
            for t in topics:
                print(f"   📚 {t['Topic']}: Pages {t['pages'][1]}-{t['pages'][2]} ({t['pages'][0]} pages)")
            
            return topics
            
        except Exception as e:
            print(f"\n⚠️ LLM topic grouping failed: {str(e)}")
            print("⚠️ Using fallback grouping...")
            return self._simple_equal_division(page_headings, total_pages)

    def _simple_equal_division(self, page_headings, total_pages):
        """Simple fallback: divide equally"""
        if total_pages <= 10:
            num_topics = 2
        elif total_pages <= 25:
            num_topics = 3
        elif total_pages <= 50:
            num_topics = 4
        else:
            num_topics = 5
        
        pages_per_topic = total_pages // num_topics
        topics = []
        current_start = 1
        
        for i in range(num_topics):
            start = current_start
            if i == num_topics - 1:
                end = total_pages
            else:
                end = start + pages_per_topic - 1
            
            count = end - start + 1
            
            # Use heading from first page of section
            topic_name = page_headings[start - 1]["heading"][:50]
            if len(topic_name) == 50:
                topic_name += "..."
            
            topics.append({
                "Topic": topic_name or f"Section {i + 1}",
                "pages": [count, start, end]
            })
            
            current_start = end + 1
        
        print(f"✅ Divided into {len(topics)} equal sections:")
        for t in topics:
            print(f"   📚 {t['Topic']}: Pages {t['pages'][1]}-{t['pages'][2]} ({t['pages'][0]} pages)")
        
        return topics

    def classify_slide_layout(self, image_path, text_content):
        """Use LLM to classify slide layout"""
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
        
        valid_layouts_str = ", ".join(self.valid_layouts)
        prompt=langfuse.get_prompt("classify",label="production")
        prompt=prompt.compile(
            valid_layouts_str=valid_layouts_str,
            text_content=text_content
        )
        
        response = self.openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a slide layout classifier. Return only the layout type."},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}}
                    ]
                }
            ],
            temperature=0.1
        )
        
        layout = response.choices[0].message.content.strip().lower()
        
        # Validate layout
        if layout not in self.valid_layouts:
            print(f"⚠️ Unknown layout '{layout}', defaulting to 'content_slide'")
            layout = "content_slide"
        
        return layout

    def add_introduction_slide(self):
        """Generate a presentation introduction page with summary of the entire content."""
        print("\n📝 Generating introduction slide...")
        
        # Collect all slide texts for global summary
        all_text = []
        for slide in self.ppt.slides:
            text = self.extract_text(slide)
            if text:
                all_text.append(text)
        combined_text = "\n".join(all_text[:200])
        prompt=langfuse.get_prompt("introduction_student",label="production")
        prompt=prompt.compile(
            combined_text=str(combined_text),
            
        )
      
        response = self.openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert instructional designer summarizing entire training presentations."
                },
                {
                    "role": "user",
                    "content": prompt
    
                }
            ],
            temperature=0.5
        )
        summary_text = response.choices[0].message.content.strip()
        
        
        intro_heading = self.doc.add_heading(level=1)
 
        self.add_formatted_paragraph(summary_text)
        
        print("✅ Introduction slide content generated and added.")

    def extract_text(self, slide):
        """Extract text from slide"""
        text = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text.append(shape.text.strip())
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            text.append("Notes: " + slide.notes_slide.notes_text_frame.text)
        return "\n".join(text)

    def generate_content_with_prompt(self, layout, image_path, text_content, previous_summary):
        """Generate handbook content using layout-specific prompt"""
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
        layout_prompt = self.prompts.get(layout, self.prompts.get("content_slide", {}))
        prompt = langfuse.get_prompt("content_student",label="production")
        full_prompt=prompt.compile(
            text_content=str(text_content),
    
        previous_summary=str(previous_summary),
        slide_specific_prompt=str(layout_prompt.get("prompt", "")
                                  ))
        
        try:
            # Use the correct API format for structured outputs
            response = self.openai_client.beta.chat.completions.parse(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are an expert instructional designer creating student handbooks."},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": full_prompt},
                            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}}
                        ]
                    }
                ],
                response_format=SlideOutput,
                temperature=0.5
            )
            
            return response.choices[0].message.parsed
            
        except Exception as e:
            print(f"⚠️ Error in content generation: {str(e)}")
            # Fallback to simple response
            return SlideOutput(
                technical_notes="Content generation failed. Please review manually.",
                # student_guidelines="• Review slide content\n• Engage learners\n• Check understanding\n• Connect to previous topics"
            )

    def generate_summary(self, content):
        """Generate 1-2 line summary of generated content"""
        if hasattr(content, "technical_notes"):
            text_to_summarize = f"{content.technical_notes}\n"
        else:
            text_to_summarize = str(content)
        
        response = self.openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Summarize the following content in 1-2 lines."},
                {"role": "user", "content": text_to_summarize}
            ],
            temperature=0.5
        )
        
        return response.choices[0].message.content.strip()

    def add_formatted_paragraph(self, text):
        """Add paragraphs with markdown-style formatting converted to Word formatting"""
        if not text:
            return
        
        # Split text into lines to handle headers
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check if line is a markdown header (### Header)
            if line.startswith('###'):
                # Remove ### and create a heading
                header_text = line.replace('###', '').strip()
                heading = self.doc.add_heading(level=3)
                run = heading.add_run(header_text)
                run.bold = True
            elif line.startswith('##'):
                # Handle ## headers (level 2)
                header_text = line.replace('##', '').strip()
                heading = self.doc.add_heading(level=2)
                run = heading.add_run(header_text)
                run.bold = True
            elif line.startswith('#'):
                # Handle # headers (level 1)
                header_text = line.replace('#', '').strip()
                heading = self.doc.add_heading(level=1)
                run = heading.add_run(header_text)
                run.bold = True
            else:
                # Regular paragraph with bold/italic formatting
                para = self.doc.add_paragraph()
                self.markdown_to_formatted_text(para, line)

    def generate_handbook(self, output_path="Enhanced_Student_Handbook.docx"):
        """Full pipeline with new architecture"""
        # Step 1: Convert PPT to PDF and images
        self.ppt_to_pdf()
        image_paths = self.pdf_to_images(self.image_dir)
        self.add_introduction_slide()
        # self.doc.add_page_break()
        image_paths = image_paths[:-1]
        total_slides = len(image_paths)
        # Step 2: Identify topics
        topics = self.identify_topics(image_paths)
        
        # Step 3: Process each topic
        for topic_idx, topic_info in enumerate(topics, start=1):
            topic_name = topic_info["Topic"]
            _, start_page, end_page = topic_info["pages"]
            

            if end_page > total_slides:
                end_page = total_slides



            print(f"\n{'='*60}")
            print(f"📚 Processing Topic {topic_idx}: {topic_name}")
            print(f"   Pages: {start_page} to {end_page}")
            print(f"{'='*60}\n")
            
            # Clear summary for new topic
            self._clear_summary()
            
            # Process each slide in topic
            for page_num in range(start_page - 1, end_page):
                if page_num >= len(self.ppt.slides):
                    break
                
                slide = self.ppt.slides[page_num]
                img_path = image_paths[page_num]
                
                print(f"🧩 Processing Slide {page_num + 1}...")
                
                # Extract text
                text = self.extract_text(slide)
                
                # Classify layout
                layout = self.classify_slide_layout(img_path, text)
                print(f"   Layout: {layout}")


                
                # Skip slides with quiz, exercise_activity, or lab_slide layouts
                if layout in self.skip_layouts:
                    print(f"   ⏭️  SKIPPING - Layout '{layout}' is in skip list")
                    continue
                # Read previous summary
                prev_summary = self._read_summary()
                
                # Generate content
                content = self.generate_content_with_prompt(layout, img_path, text, prev_summary)
                
                # Generate and save summary
                summary = self.generate_summary(content)
                self._append_summary(f"Slide {page_num + 1}: {summary}")
                
                # Add to document
                self.doc.add_page_break()
                
                # Add slide image
                self.doc.add_picture(img_path, width=Inches(6.5))
                
                # Add Technical Notes
                self.doc.add_heading("Notes", level=2)
                self.add_formatted_paragraph(content.technical_notes)
                
                # Add Trainer Guidelines
                # self.doc.add_heading("Student Guidelines", level=2)
                # self.add_formatted_paragraph(content.student_guidelines)
                
                # self.doc.add_page_break()
        
        # Save document
        self.doc.save(output_path)
        print(f"\n✅ Handbook generated successfully: {output_path}")

# Usage
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Lifespan context manager for startup and shutdown"""
    # Startup
    cleanup_scheduler.start()
    print("✅ Auto-cleanup scheduler started - will run every 60 minutes")
    print(f"⏰ Jobs older than 2 hours will be automatically deleted")
    print(f"📊 Found {len(job_manager.list_jobs())} existing jobs")
    
    yield  # Application runs here
    
    # Shutdown
    cleanup_scheduler.stop()
    print("✅ Auto-cleanup scheduler stopped")

# STEP 6: Update FastAPI app initialization
app = FastAPI(
    title="AI Handbook Generator API",
    description="Generate trainer handbooks from PowerPoint presentations",
    version="1.0.0",
    lifespan=lifespan  # Add this parameter
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

def process_handbook_background(job_id: str, ppt_path: str, api_key: str):
    """Background task to process the handbook with proper isolation"""
    try:
        # Update status
        job_manager.update_job(job_id, {
            "status": "processing",
            "progress": "Initializing..."
        })
        
        # Create isolated workspace for this job
        job_workspace = UPLOAD_DIR / job_id
        
        # Import here to avoid circular dependencies
        from ppt2hb_student import AIHandbookMakerStudent  # Replace with actual import
        
        # Initialize with job-specific paths
        maker = AIHandbookMakerStudent(ppt_path, api_key=api_key)
        
        # Override base_dir to use job-specific workspace
        maker.base_dir = str(job_workspace)
        maker.pdf_path = str(job_workspace / "slides.pdf")
        maker.image_dir = str(job_workspace / "slide_images")
        maker.summary_path = str(job_workspace / "summary.txt")
        
        os.makedirs(maker.image_dir, exist_ok=True)
        
        job_manager.update_job(job_id, {
            "progress": "Converting presentation to PDF..."
        })
        
        # Generate unique output filename
        output_filename = f"handbook_{job_id}.docx"
        output_path = OUTPUT_DIR / output_filename
        
        job_manager.update_job(job_id, {
            "progress": "Generating handbook content..."
        })
        
        maker.generate_handbook(str(output_path))
        
        # Success
        job_manager.update_job(job_id, {
            "status": "completed",
            "completed_at": datetime.now().isoformat(),
            "output_file": output_filename,
            "progress": "Completed successfully"
        })
        
    except Exception as e:
        error_msg = str(e)
        print(f"Error processing job {job_id}: {error_msg}")
        
        job_manager.update_job(job_id, {
            "status": "failed",
            "error": error_msg,
            "completed_at": datetime.now().isoformat()
        })
    finally:
        # Clean up workspace files (keep only output)
        try:
            job_workspace = UPLOAD_DIR / job_id
            for item in job_workspace.iterdir():
                if item.is_file() and item.name != Path(ppt_path).name:
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
        except Exception as e:
            print(f"Error cleaning workspace for job {job_id}: {str(e)}")
@app.get("/admin/debug/jobs")
async def debug_jobs():
    """Debug: Check all jobs"""
    from datetime import datetime, timezone
    
    jobs = job_manager.list_jobs()
    now = datetime.now(timezone.utc)
    
    jobs_info = []
    for job in jobs:
        created_at_str = job.get('created_at')
        try:
            created_at = datetime.fromisoformat(created_at_str.replace('Z', '+00:00'))
            if created_at.tzinfo is None:
                created_at = created_at.replace(tzinfo=timezone.utc)
            
            age = now - created_at
            age_hours = age.total_seconds() / 3600
            
            jobs_info.append({
                'job_id': job.get('job_id'),
                'status': job.get('status'),
                'created_at': created_at_str,
                'age_hours': round(age_hours, 2),
                'will_delete': age_hours > 2.0  # Based on 2 hour lifetime
            })
        except Exception as e:
            jobs_info.append({
                'job_id': job.get('job_id'),
                'error': str(e),
                'created_at': created_at_str
            })
    
    return {
        'current_time': now.isoformat(),
        'total_jobs': len(jobs),
        'jobs': jobs_info,
        'storage_file': job_manager.storage_file
    }

@app.get("/admin/debug/files")
async def debug_files():
    """Debug: List all files"""
    from pathlib import Path
    
    uploads_dir = Path("uploads")
    outputs_dir = Path("outputs")
    
    upload_dirs = []
    if uploads_dir.exists():
        upload_dirs = [d.name for d in uploads_dir.iterdir() if d.is_dir()]
    
    output_files = []
    if outputs_dir.exists():
        output_files = [f.name for f in outputs_dir.iterdir() if f.is_file()]
    
    return {
        'upload_directories': upload_dirs,
        'output_files': output_files,
        'upload_dir_count': len(upload_dirs),
        'output_file_count': len(output_files)
    }

@app.post("/admin/cleanup/force")
async def force_cleanup():
    """Force immediate cleanup"""
    stats = cleanup_scheduler.force_cleanup_now()
    return {
        "message": "Cleanup completed",
        "stats": stats
    }

@app.get("/admin/cleanup/status")
async def cleanup_status():
    """Check cleanup scheduler status"""
    return {
        "running": cleanup_scheduler.is_running(),
        "interval_minutes": cleanup_scheduler.cleanup_interval / 60,
        "job_lifetime_hours": cleanup_scheduler.job_lifetime.total_seconds() / 3600,
        "total_jobs": len(job_manager.list_jobs())
    }

@app.post("/admin/cleanup/test")
async def test_cleanup_short_lifetime():
    """Test cleanup with 1 minute lifetime (for testing only)"""
    from cleanup import AutoCleanupScheduler
    
    test_scheduler = AutoCleanupScheduler(
        job_manager=job_manager,
        job_lifetime_hours=0.0167,  # 1 minute
        upload_dir="uploads",
        output_dir="outputs"
    )
    
    stats = test_scheduler.cleanup_old_jobs()
    return {
        "message": "Test cleanup with 1-minute lifetime completed",
        "stats": stats
    }
@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "message": "AI Handbook Generator API",
        "version": "1.0.0",
        "active_jobs": job_manager.get_active_count(),
        "endpoints": {
            "POST /upload": "Upload PowerPoint",
            "GET /status/{job_id}": "Check status",
            "GET /download/{job_id}": "Download handbook",
            "GET /jobs": "List all jobs",
            "DELETE /job/{job_id}": "Delete job"
        }
    }

@app.post("/upload", response_model=JobResponse)
async def upload_presentation(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    api_key: str = Form(...)
):
    """Upload PowerPoint and start processing"""
    # Validate file type
    if not file.filename.endswith(('.ppt', '.pptx')):
        raise HTTPException(
            status_code=400, 
            detail="Only PowerPoint files (.ppt, .pptx) are accepted"
        )
    
    # Validate API key
    if not api_key or len(api_key) < 10:
        raise HTTPException(
            status_code=400,
            detail="Valid OpenAI API key required"
        )
    
    # Generate unique job ID
    job_id = str(uuid.uuid4())
    
    # Create isolated job directory
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    
    # Save uploaded file with original name
    file_path = job_dir / file.filename
    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        # Cleanup on failure
        if job_dir.exists():
            shutil.rmtree(job_dir)
        raise HTTPException(
            status_code=500, 
            detail=f"Failed to save file: {str(e)}"
        )
    
    # Create job record
    job_data = {
        "job_id": job_id,
        "status": "pending",
        "progress": "File uploaded successfully",
        "created_at": datetime.now().isoformat(),
        "completed_at": None,
        "error": None,
        "output_file": None,
        "original_filename": file.filename
    }
    
    job_manager.create_job(job_id, job_data)
    
    # Start background processing
    background_tasks.add_task(
        process_handbook_background, 
        job_id, 
        str(file_path), 
        api_key
    )
    
    return JobResponse(
        job_id=job_id,
        message="File uploaded successfully. Processing started.",
        status="pending"
    )
def cleanup_job_files(job_id: str):
    """Clean up all files associated with a job"""
    try:
        # Remove upload directory
        job_dir = UPLOAD_DIR / job_id
        if job_dir.exists():
            shutil.rmtree(job_dir)
            print(f"✅ Deleted upload directory: {job_dir}")
        
        # Remove output file
        output_file = OUTPUT_DIR / f"handbook_{job_id}.docx"
        if output_file.exists():
            output_file.unlink()
            print(f"✅ Deleted output file: {output_file.name}")
            
    except Exception as e:
        print(f"❌ Error cleaning up job {job_id}: {str(e)}")
@app.get("/status/{job_id}", response_model=JobStatus)
async def get_job_status(job_id: str):
    """Get job status"""
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return JobStatus(**job)

@app.get("/download/{job_id}")
async def download_handbook(job_id: str):
    """Download generated handbook"""
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    if job["status"] != "completed":
        raise HTTPException(
            status_code=400, 
            detail=f"Job status: {job['status']}. {job.get('progress', '')}"
        )
    
    output_file = OUTPUT_DIR / job["output_file"]
    if not output_file.exists():
        raise HTTPException(status_code=404, detail="Output file not found")
    
    return FileResponse(
        path=output_file,
        filename=f"Enhanced_Trainer_Handbook.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/jobs")
async def list_jobs():
    """List all jobs"""
    jobs = job_manager.list_jobs()
    return {
        "total_jobs": len(jobs),
        "active_jobs": job_manager.get_active_count(),
        "jobs": jobs
    }

@app.delete("/job/{job_id}")
async def delete_job(job_id: str):
    """Delete a job and its associated files"""
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    # Clean up files
    cleanup_job_files(job_id)
    
    # Remove from job manager
    job_manager.delete_job(job_id)
    
    return {"message": f"Job {job_id} deleted successfully"}

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "active_jobs": job_manager.get_active_count(),
        "total_jobs": len(job_manager.list_jobs())
    }

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8017)
