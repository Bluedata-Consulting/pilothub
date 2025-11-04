import streamlit as st
import requests
import time
from pathlib import Path
import json

# Configuration - Multiple API endpoints
API_ENDPOINTS = {
    "student": "http://127.0.0.1:8017",
    "trainer": "http://127.0.0.1:8016"
}

# Page config
st.set_page_config(
    page_title="Handbook Generator",
    page_icon="📚",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .status-box {
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .status-queued, .status-pending {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
    }
    .status-processing {
        background-color: #cfe2ff;
        border-left: 4px solid #0d6efd;
    }
    .status-completed {
        background-color: #d1e7dd;
        border-left: 4px solid #198754;
    }
    .status-failed {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
    }
    .handbook-card {
        border: 2px solid #e0e0e0;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        background: white;
    }
    .handbook-card:hover {
        border-color: #1f77b4;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .api-status {
        padding: 10px;
        border-radius: 5px;
        margin: 5px 0;
    }
    .api-online {
        background-color: #d1e7dd;
        border-left: 4px solid #198754;
    }
    .api-offline {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
def init_session_state():
    """Initialize all session state variables"""
    defaults = {
        'job_id': None,
        'job_status': None,
        'auto_refresh': False,
        'api_key': "",
        'selected_handbook_type': "trainer",
        'show_upload_form': True,
        'last_refresh_time': 0,
        'refresh_count': 0
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# Helper functions
def clear_job_cache(job_id):
    """Clear all cached data for a specific job"""
    cache_key = f"handbook_data_{job_id}"
    if cache_key in st.session_state:
        del st.session_state[cache_key]

def reset_job_state():
    """Reset all job-related session state"""
    st.session_state.job_id = None
    st.session_state.job_status = None
    st.session_state.auto_refresh = False
    st.session_state.show_upload_form = True
    st.session_state.refresh_count = 0
    st.session_state.last_refresh_time = 0
    
    # Clear any handbook data caches
    keys_to_delete = [key for key in st.session_state.keys() if key.startswith('handbook_data_')]
    for key in keys_to_delete:
        del st.session_state[key]

def get_cached_handbook(job_id, handbook_type):
    """Get handbook data with caching"""
    cache_key = f"handbook_data_{job_id}"
    
    # Return cached data if available
    if cache_key in st.session_state:
        return st.session_state[cache_key]
    
    # Otherwise fetch and cache
    handbook_data = download_handbook(job_id, handbook_type)
    if handbook_data:
        st.session_state[cache_key] = handbook_data
    
    return handbook_data

def check_api_health(handbook_type):
    """Check if API is running"""
    try:
        api_url = API_ENDPOINTS.get(handbook_type)
        response = requests.get(f"{api_url}/health", timeout=5)
        if response.status_code == 200:
            return True, response.json()
        return False, None
    except Exception as e:
        return False, str(e)

def get_api_url(handbook_type):
    """Get API URL based on handbook type"""
    return API_ENDPOINTS.get(handbook_type, API_ENDPOINTS["trainer"])

def upload_file(file, api_key, handbook_type):
    """Upload file to API"""
    try:
        api_url = get_api_url(handbook_type)
        
        # Reset file pointer to beginning
        file.seek(0)
        
        files = {
            'file': (file.name, file, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
        }
        data = {'api_key': api_key}
        
        response = requests.post(
            f"{api_url}/upload", 
            files=files, 
            data=data,
            timeout=60  # Increased timeout for large files
        )
        
        if response.status_code == 200:
            return response.json()
        else:
            try:
                error_detail = response.json().get('detail', response.text)
            except:
                error_detail = response.text
            st.error(f"Upload failed (Status {response.status_code}): {error_detail}")
            return None
            
    except requests.exceptions.Timeout:
        st.error("Upload request timed out. Please check your file size and try again.")
        return None
    except Exception as e:
        st.error(f"Error uploading file: {str(e)}")
        return None

def get_job_status(job_id, handbook_type):
    """Get job status from API"""
    try:
        api_url = get_api_url(handbook_type)
        response = requests.get(f"{api_url}/status/{job_id}", timeout=10)
        
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 404:
            st.warning("Job not found. It may have been cleaned up or expired.")
            return None
        else:
            st.error(f"Error getting status (Status {response.status_code})")
            return None
            
    except requests.exceptions.Timeout:
        st.warning("Status check timed out. Will retry...")
        return None
    except Exception as e:
        st.error(f"Error getting status: {str(e)}")
        return None

def download_handbook(job_id, handbook_type):
    """Download handbook from API with better error handling"""
    try:
        api_url = get_api_url(handbook_type)
        response = requests.get(
            f"{api_url}/download/{job_id}", 
            timeout=60,
            stream=True
        )
        
        if response.status_code == 200:
            return response.content
        elif response.status_code == 404:
            st.error("Handbook file not found. The job may not be completed yet.")
            return None
        else:
            try:
                error_detail = response.json().get('detail', response.text)
            except:
                error_detail = response.text
            st.error(f"Download failed (Status {response.status_code}): {error_detail}")
            return None
            
    except requests.exceptions.Timeout:
        st.error("Download request timed out. Please try again.")
        return None
    except Exception as e:
        st.error(f"Error downloading handbook: {str(e)}")
        return None

def delete_job(job_id, handbook_type):
    """Delete a job and clean up resources"""
    try:
        api_url = get_api_url(handbook_type)
        response = requests.delete(f"{api_url}/job/{job_id}", timeout=10)
        
        if response.status_code == 200:
            return True
        else:
            st.warning(f"Cleanup returned status {response.status_code}")
            return False
            
    except Exception as e:
        st.error(f"Error during cleanup: {str(e)}")
        return False

# Header
st.markdown('<div class="main-header">📚 Handbook Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header"></div>', unsafe_allow_html=True)

# Sidebar for API Key and Settings
with st.sidebar:
    st.header("🔑 Configuration")
    
    api_key_input = st.text_input(
        "OpenAI API Key",
        type="password",
        value=st.session_state.api_key,
        help="Enter your OpenAI API key. It will be stored in your session.",
        placeholder="sk-..."
    )
    
    if api_key_input:
        if len(api_key_input) >= 20:  # Basic validation
            st.session_state.api_key = api_key_input
            st.success("✅ API Key saved")
        else:
            st.warning("⚠️ API key seems too short")
    else:
        st.warning("⚠️ Please enter your OpenAI API key")
    
    st.markdown("---")
    
    # Settings
    st.header("⚙️ Settings")
    
    auto_refresh_enabled = st.checkbox(
        "Auto-refresh status",
        value=True,
        help="Automatically refresh job status while processing"
    )
    
    if auto_refresh_enabled != st.session_state.auto_refresh and st.session_state.job_id:
        st.session_state.auto_refresh = auto_refresh_enabled
    
    st.markdown("---")
    
    # st.markdown("### 📖 About")
    # st.markdown("""
    # This tool uses AI to analyze your presentation and generate comprehensive handbooks.
    
    # **Features:**
    # - 👨‍🎓 Student handbooks
    # - 👨‍🏫 Trainer guides
    # - 🤖 AI-powered content
    # - 📊 Progress tracking
    # """)
    
    st.markdown("---")
    
    # Show current job info if exists
    if st.session_state.job_id:
        st.info(f"**Active Job ID:**\n`{st.session_state.job_id[:16]}...`")
        st.caption(f"Type: {st.session_state.selected_handbook_type.capitalize()}")

# Check API health for both services
st.markdown("### 🔌 API Health Status")
col1, col2 = st.columns(2)

student_status = None
trainer_status = None

with col1:
    student_api_online, student_status = check_api_health("student")
    status_class = "api-online" if student_api_online else "api-offline"
    status_icon = "✅" if student_api_online else "❌"
    st.markdown(f"""
        <div class="api-status {status_class}">
            <strong>{status_icon} Student API :</strong> {'Online' if student_api_online else 'Offline'}
        </div>
    """, unsafe_allow_html=True)
    
    if student_api_online and student_status:
        with st.expander("📊 Student API Stats"):
            st.json({
                "Active Jobs": student_status.get("active_jobs", 0),
                "Total Jobs": student_status.get("total_jobs", 0),
                "Status": "Healthy"
            })

with col2:
    trainer_api_online, trainer_status = check_api_health("trainer")
    status_class = "api-online" if trainer_api_online else "api-offline"
    status_icon = "✅" if trainer_api_online else "❌"
    st.markdown(f"""
        <div class="api-status {status_class}">
            <strong>{status_icon} Trainer API :</strong> {'Online' if trainer_api_online else 'Offline'}
        </div>
    """, unsafe_allow_html=True)
    
    if trainer_api_online and trainer_status:
        with st.expander("📊 Trainer API Stats"):
            st.json({
                "Active Jobs": trainer_status.get("active_jobs", 0),
                "Total Jobs": trainer_status.get("total_jobs", 0),
                "Status": "Healthy"
            })

if not student_api_online and not trainer_api_online:
    st.error("⚠️ Both APIs are offline! Please start at least one FastAPI server.")
    st.code("""
# For Student Handbook:
python student_app.py
# or
uvicorn student_app:app --port 8080

# For Trainer Handbook:
python trainer_app.py
# or
uvicorn trainer_app:app --port 8016
    """, language="bash")
    st.stop()

st.markdown("---")

# Auto-refresh logic for active jobs
if st.session_state.job_id and st.session_state.auto_refresh:
    current_time = time.time()
    
    # Throttle refresh to every 2 seconds
    if current_time - st.session_state.last_refresh_time > 2:
        status = get_job_status(st.session_state.job_id, st.session_state.selected_handbook_type)
        
        if status:
            st.session_state.job_status = status
            st.session_state.last_refresh_time = current_time
            st.session_state.refresh_count += 1
            
            # Stop auto-refresh if job is complete or failed
            if status['status'] not in ['processing', 'queued', 'pending']:
                st.session_state.auto_refresh = False
            else:
                # Continue auto-refresh
                time.sleep(2)
                st.rerun()

# Main content - Show either upload form or job status
if st.session_state.show_upload_form and not st.session_state.job_id:
    # ===== UPLOAD SECTION =====
    st.header("📤 Upload & Generate Handbook")
    
    if not st.session_state.api_key or len(st.session_state.api_key) < 20:
        st.warning("⚠️ Please enter a valid OpenAI API Key in the sidebar first!")
        st.stop()
    
    # Handbook Type Selection
    st.subheader("1️⃣ Select Handbook Type")
    
    # col1, col2 = st.columns(2)
    
    # with col1:
    #     st.markdown("""
    #     <div class="handbook-card">
    #         <h3>👨‍🎓 Student Handbook</h3>
    #         <p><strong>Perfect for:</strong></p>
    #         <ul>
    #             <li>Student study materials</li>
    #             <li>Self-paced learning</li>
    #             <li>Reference documentation</li>
    #         </ul>
    #         <p><strong>API Port:</strong> 8080</p>
    #     </div>
    #     """, unsafe_allow_html=True)
    
    # with col2:
    #     st.markdown("""
    #     <div class="handbook-card">
    #         <h3>👨‍🏫 Trainer Handbook</h3>
    #         <p><strong>Perfect for:</strong></p>
    #         <ul>
    #             <li>Instructor guides</li>
    #             <li>Training delivery</li>
    #             <li>Teaching strategies</li>
    #         </ul>
    #         <p><strong>API Port:</strong> 8016</p>
    #     </div>
    #     """, unsafe_allow_html=True)
    
    handbook_type = st.radio(
        "Choose handbook type:",
        options=["student", "trainer"],
        format_func=lambda x: "👨‍🎓 Student Handbook" if x == "student" else "👨‍🏫 Trainer Handbook",
        horizontal=True,
        index=1 if st.session_state.selected_handbook_type == "trainer" else 0
    )
    
    st.session_state.selected_handbook_type = handbook_type
    
    # Check if selected API is available
    api_available = check_api_health(handbook_type)[0]
    
    if not api_available:
        st.error(f"⚠️ {handbook_type.capitalize()} API is not running!")
        st.info(f"Please start the {handbook_type} handbook API server on port {8016}.")
        st.stop()
    else:
        st.success(f"✅ {handbook_type.capitalize()} API is ready")
    
    st.markdown("---")
    
    # File Upload
    st.subheader("2️⃣ Upload Your Presentation")
    
    uploaded_file = st.file_uploader(
        "Choose a PowerPoint file",
        type=['ppt', 'pptx'],
        help="Upload your PowerPoint presentation (.ppt or .pptx)"
    )
    
    if uploaded_file:
        # Show file info
        file_size_mb = uploaded_file.size / (1024 * 1024)
        st.info(f"📄 **File:** {uploaded_file.name}")
        st.info(f"📊 **Size:** {file_size_mb:.2f} MB")
        
        # Warn for large files
        if file_size_mb > 50:
            st.warning("⚠️ Large file detected. Processing may take longer.")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            if st.button("🚀 Generate Handbook", type="primary", use_container_width=True):
                
                # Validate API key length
                if len(st.session_state.api_key) < 20:
                    st.error("Invalid API key. Please check and try again.")
                    st.stop()
                
                with st.spinner(f"Uploading {uploaded_file.name}..."):
                    result = upload_file(uploaded_file, st.session_state.api_key, handbook_type)
                    
                    if result and 'job_id' in result:
                        st.session_state.job_id = result['job_id']
                        st.session_state.auto_refresh = True
                        st.session_state.show_upload_form = False
                        st.session_state.refresh_count = 0
                        
                        st.success("✅ Upload successful!")
                        st.success(f"🔑 Job ID: `{result['job_id']}`")
                        st.success(f"📚 Generating **{handbook_type.capitalize()} Handbook**")
                        
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error("Upload failed. Please check the error message above and try again.")

else:
    # ===== JOB STATUS SECTION =====
    st.header("📊 Job Status & Download")
    
    if st.session_state.job_id:
        job_id_display = st.session_state.job_id
        status_handbook_type = st.session_state.selected_handbook_type
        
        # Show job info
        col1, col2 = st.columns([3, 1])
        with col1:
            st.info(f"**Job ID:** `{job_id_display}`")
        with col2:
            st.caption(f"Type: {status_handbook_type.capitalize()}")
        
        # Action buttons
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                st.session_state.job_status = get_job_status(job_id_display, status_handbook_type)
                st.rerun()
        
        with col2:
            auto_refresh_btn = st.checkbox(
                "Auto-refresh",
                value=st.session_state.auto_refresh,
                key="auto_refresh_toggle"
            )
            if auto_refresh_btn != st.session_state.auto_refresh:
                st.session_state.auto_refresh = auto_refresh_btn
                st.rerun()
        
        with col3:
            if st.button("📤 New Upload", use_container_width=True):
                reset_job_state()
                st.rerun()
        
        with col4:
            if st.button("🗑️ Delete Job", use_container_width=True):
                if delete_job(job_id_display, status_handbook_type):
                    st.success("Job deleted!")
                    reset_job_state()
                    time.sleep(1)
                    st.rerun()
        
        st.markdown("---")
        
        # Get and display status
        status = st.session_state.job_status if st.session_state.job_status else get_job_status(job_id_display, status_handbook_type)
        
        if status:
            st.session_state.job_status = status
            
            # Status display
            status_emoji = {
                'queued': '⏳',
                'pending': '⏳',
                'processing': '⚙️',
                'completed': '✅',
                'failed': '❌'
            }
            
            current_status = status.get('status', 'unknown')
            status_class = f"status-{current_status}"
            
            st.markdown(f"""
                <div class="status-box {status_class}">
                    <h3>{status_emoji.get(current_status, '📊')} Status: {current_status.upper()}</h3>
                    <p><strong>Progress:</strong> {status.get('progress', 'Initializing...')}</p>
                </div>
            """, unsafe_allow_html=True)
            
            # Show progress bar for processing
            if current_status == 'processing':
                # Indeterminate progress bar
                st.progress(0.5, text="Processing slides...")
                
                if st.session_state.auto_refresh:
                    st.caption(f"🔄 Auto-refreshing... (refresh #{st.session_state.refresh_count})")
            
            elif current_status in ['queued', 'pending']:
                st.progress(0.0, text="Waiting in queue...")
                
                if st.session_state.auto_refresh:
                    st.caption(f"🔄 Auto-refreshing... (refresh #{st.session_state.refresh_count})")
            
            elif current_status == 'completed':
                st.progress(1.0, text="✅ Completed!")
                
                st.success(f"🎉 Your {status_handbook_type.capitalize()} Handbook is ready!")
                
                # Download section
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    handbook_data = get_cached_handbook(job_id_display, status_handbook_type)
                    
                    if handbook_data:
                        st.download_button(
                            label=f"📥 Download {status_handbook_type.capitalize()} Handbook",
                            data=handbook_data,
                            file_name=f"Enhanced_{status_handbook_type.capitalize()}_Handbook.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )
                        st.caption(f"📊 File size: {len(handbook_data) / 1024:.1f} KB")
                    else:
                        st.error("❌ Failed to retrieve handbook. Please try refreshing.")
                        if st.button("🔄 Retry Download"):
                            clear_job_cache(job_id_display)
                            st.rerun()
                
                with col2:
                    st.metric("Status", "Done ✅")
                    if status.get('completed_at'):
                        st.caption(f"Completed at:\n{status['completed_at'][:19]}")
            
            elif current_status == 'failed':
                st.error("❌ Processing failed")
                
                error_msg = status.get('error', 'Unknown error occurred')
                
                with st.expander("📋 Error Details", expanded=True):
                    st.code(error_msg)
                
                st.warning("💡 **Troubleshooting:**")
                st.markdown("""
                - Check if your PowerPoint file is valid
                - Ensure your API key is correct
                - Try uploading a different file
                - Check the API logs for more details
                """)
                
                if st.button("🔄 Try Again with New File"):
                    reset_job_state()
                    st.rerun()
        
        else:
            st.warning("⚠️ Could not retrieve job status. The job may have been deleted or expired.")
            
            if st.button("🔙 Start Over"):
                reset_job_state()
                st.rerun()
    
    else:
        st.info("ℹ️ No active job. Upload a presentation to get started.")
        
        if st.button("📤 Upload New File", use_container_width=True, type="primary"):
            reset_job_state()
            st.rerun()

# Footer
st.markdown("---")
# st.markdown(
#     '<div style="text-align: center; color: #666; font-size: 0.9rem;">'
#     'AI Handbook Generator v2.1 | Multi-User Support | Student (8080) & Trainer (8016) | Powered by OpenAI GPT-4'
#     '</div>',
#     unsafe_allow_html=True
# )