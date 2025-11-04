import shutil
import time
import threading
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import Optional
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class AutoCleanupScheduler:
    """
    Automatic cleanup scheduler for jobs and associated files.
    Runs cleanup every hour and deletes jobs older than 1 hour.
    """
    
    def __init__(
        self,
        job_manager,
        cleanup_interval_minutes: int = 60,
        job_lifetime_hours: int = 1,
        upload_dir: str = "uploads",
        output_dir: str = "outputs"
    ):
        """
        Initialize the cleanup scheduler.
        
        Args:
            job_manager: The JobManager instance
            cleanup_interval_minutes: How often to run cleanup (default: 60 minutes)
            job_lifetime_hours: How long jobs should live (default: 1 hour)
            upload_dir: Directory where uploads are stored
            output_dir: Directory where output files are stored
        """
        self.job_manager = job_manager
        self.cleanup_interval = cleanup_interval_minutes * 60  # Convert to seconds
        self.job_lifetime = timedelta(hours=job_lifetime_hours)
        self.upload_dir = Path(upload_dir)
        self.output_dir = Path(output_dir)
        
        self._stop_event = threading.Event()
        self._cleanup_thread: Optional[threading.Thread] = None
        self._is_running = False
        
        logger.info(f"AutoCleanupScheduler initialized:")
        logger.info(f"  - Cleanup interval: {cleanup_interval_minutes} minutes")
        logger.info(f"  - Job lifetime: {job_lifetime_hours} hour(s)")
        logger.info(f"  - Upload directory: {self.upload_dir}")
        logger.info(f"  - Output directory: {self.output_dir}")
    
    def _cleanup_job_files(self, job_id: str):
        """
        Clean up all files associated with a job.
        
        Args:
            job_id: The job ID to clean up
        """
        files_deleted = []
        
        # Clean upload directory
        job_upload_dir = self.upload_dir / job_id
        if job_upload_dir.exists():
            try:
                shutil.rmtree(job_upload_dir)
                files_deleted.append(f"upload_dir: {job_upload_dir}")
                logger.debug(f"Deleted upload directory: {job_upload_dir}")
            except Exception as e:
                logger.error(f"Error deleting upload directory {job_upload_dir}: {e}")
        
        # Clean output file
        output_file = self.output_dir / f"handbook_{job_id}.docx"
        if output_file.exists():
            try:
                output_file.unlink()
                files_deleted.append(f"handbook: {output_file.name}")
                logger.debug(f"Deleted output file: {output_file}")
            except Exception as e:
                logger.error(f"Error deleting output file {output_file}: {e}")
        
        return files_deleted
    
    def _parse_datetime(self, dt_str: str) -> datetime:
        """
        Parse datetime string and ensure it's timezone-aware.
        
        Args:
            dt_str: ISO format datetime string
            
        Returns:
            Timezone-aware datetime object
        """
        dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
        
        # If naive, assume UTC
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        
        return dt
    
    def _should_delete_job(self, job_data: dict) -> bool:
        """
        Check if a job should be deleted based on its age.
        
        Args:
            job_data: Job data dictionary
            
        Returns:
            True if job should be deleted, False otherwise
        """
        try:
            created_at_str = job_data.get('created_at')
            if not created_at_str:
                logger.warning(f"Job missing created_at field: {job_data.get('job_id')}")
                return False
            
            # Parse the created_at timestamp
            created_at = self._parse_datetime(created_at_str)
            now = datetime.now(timezone.utc)
            age = now - created_at
            
            job_id = job_data.get('job_id', 'unknown')
            should_delete = age > self.job_lifetime
            
            # Debug logging
            logger.debug(f"Job {job_id}: created={created_at}, age={age}, lifetime={self.job_lifetime}, delete={should_delete}")
            
            return should_delete
            
        except Exception as e:
            logger.error(f"Error checking job age for {job_data.get('job_id')}: {e}")
            return False
    
    def cleanup_old_jobs(self) -> dict:
        """
        Perform cleanup of old jobs.
        
        Returns:
            Dictionary with cleanup statistics
        """
        logger.info("🧹 Starting cleanup of old jobs...")
        
        stats = {
            'jobs_checked': 0,
            'jobs_deleted': 0,
            'files_deleted': 0,
            'errors': 0,
            'start_time': datetime.now(timezone.utc).isoformat()
        }
        
        try:
            # Get all jobs
            all_jobs = self.job_manager.list_jobs()
            stats['jobs_checked'] = len(all_jobs)
            
            logger.info(f"Checking {len(all_jobs)} jobs for cleanup...")
            
            cutoff_time = datetime.now(timezone.utc) - self.job_lifetime
            logger.info(f"Cutoff time: {cutoff_time.isoformat()} (jobs older than this will be deleted)")
            
            for job_data in all_jobs:
                job_id = job_data.get('job_id')
                
                if not job_id:
                    logger.warning("Found job without job_id, skipping")
                    stats['errors'] += 1
                    continue
                
                # Check if job should be deleted
                if self._should_delete_job(job_data):
                    created_at = job_data.get('created_at', 'unknown')
                    status = job_data.get('status', 'unknown')
                    
                    logger.info(f"Deleting old job: {job_id}")
                    logger.info(f"  - Created: {created_at}")
                    logger.info(f"  - Status: {status}")
                    
                    try:
                        # Clean up files
                        files_deleted = self._cleanup_job_files(job_id)
                        stats['files_deleted'] += len(files_deleted)
                        
                        # Delete job from manager
                        self.job_manager.delete_job(job_id)
                        stats['jobs_deleted'] += 1
                        
                        logger.info(f"✅ Deleted job {job_id} and {len(files_deleted)} associated files")
                        
                    except Exception as e:
                        logger.error(f"❌ Error deleting job {job_id}: {e}")
                        stats['errors'] += 1
                else:
                    # Log why job wasn't deleted
                    created_at = job_data.get('created_at', 'unknown')
                    logger.debug(f"Keeping job {job_id} (created: {created_at})")
            
            # Summary
            stats['end_time'] = datetime.now(timezone.utc).isoformat()
            
            logger.info(f"🎉 Cleanup completed!")
            logger.info(f"  - Jobs checked: {stats['jobs_checked']}")
            logger.info(f"  - Jobs deleted: {stats['jobs_deleted']}")
            logger.info(f"  - Files deleted: {stats['files_deleted']}")
            logger.info(f"  - Errors: {stats['errors']}")
            
            return stats
            
        except Exception as e:
            logger.error(f"❌ Critical error during cleanup: {e}", exc_info=True)
            stats['errors'] += 1
            stats['end_time'] = datetime.now(timezone.utc).isoformat()
            return stats
    
    def _cleanup_loop(self):
        """
        Main cleanup loop that runs in a separate thread.
        """
        logger.info(f"🚀 Cleanup scheduler started (interval: {self.cleanup_interval/60:.0f} minutes)")
        
        while not self._stop_event.is_set():
            try:
                # Perform cleanup
                self.cleanup_old_jobs()
                
                # Wait for next interval (or until stopped)
                logger.info(f"⏰ Next cleanup in {self.cleanup_interval/60:.0f} minutes...")
                self._stop_event.wait(self.cleanup_interval)
                
            except Exception as e:
                logger.error(f"Error in cleanup loop: {e}", exc_info=True)
                # Wait a bit before retrying
                time.sleep(60)
        
        logger.info("🛑 Cleanup scheduler stopped")
    
    def start(self):
        """
        Start the automatic cleanup scheduler.
        """
        if self._is_running:
            logger.warning("Cleanup scheduler is already running")
            return
        
        self._stop_event.clear()
        self._cleanup_thread = threading.Thread(
            target=self._cleanup_loop,
            name="CleanupScheduler",
            daemon=True
        )
        self._cleanup_thread.start()
        self._is_running = True
        
        logger.info("✅ Cleanup scheduler started successfully")
    
    def stop(self):
        """
        Stop the automatic cleanup scheduler.
        """
        if not self._is_running:
            logger.warning("Cleanup scheduler is not running")
            return
        
        logger.info("Stopping cleanup scheduler...")
        self._stop_event.set()
        
        if self._cleanup_thread:
            self._cleanup_thread.join(timeout=5)
        
        self._is_running = False
        logger.info("✅ Cleanup scheduler stopped")
    
    def is_running(self) -> bool:
        """
        Check if the cleanup scheduler is running.
        
        Returns:
            True if running, False otherwise
        """
        return self._is_running
    
    def force_cleanup_now(self) -> dict:
        """
        Force an immediate cleanup (useful for testing or manual cleanup).
        
        Returns:
            Dictionary with cleanup statistics
        """
        logger.info("🔥 Forcing immediate cleanup...")
        return self.cleanup_old_jobs()


# Convenience function for one-time cleanup
def cleanup_old_jobs_once(
    job_manager,
    job_lifetime_hours: int = 1,
    upload_dir: str = "uploads",
    output_dir: str = "outputs"
) -> dict:
    """
    Perform a one-time cleanup of old jobs.
    
    Args:
        job_manager: The JobManager instance
        job_lifetime_hours: How long jobs should live (default: 1 hour)
        upload_dir: Directory where uploads are stored
        output_dir: Directory where output files are stored
        
    Returns:
        Dictionary with cleanup statistics
    """
    scheduler = AutoCleanupScheduler(
        job_manager=job_manager,
        job_lifetime_hours=job_lifetime_hours,
        upload_dir=upload_dir,
        output_dir=output_dir
    )
    return scheduler.cleanup_old_jobs()