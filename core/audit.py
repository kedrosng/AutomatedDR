"""
Audit module for JV Cost Control System
Manages SHA3-256 snapshots and change tracking
"""

import hashlib
import json
import pandas as pd
from datetime import datetime
import os
from typing import Dict, Any, List
import logging

logger = logging.getLogger(__name__)

def save_snapshot(data: pd.DataFrame, discrepancies: List[Dict[str, Any]], 
                 snapshot_dir: str = "data/archive/") -> str:
    """
    Create and save a SHA3-256 hash snapshot of the current data state
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot_subdir = os.path.join(snapshot_dir, f"snapshots/{timestamp}")
    
    os.makedirs(snapshot_subdir, exist_ok=True)
    
    # Convert data to JSON for hashing
    data_json = data.to_json(orient='records', date_format='iso')
    discrepancies_json = json.dumps(discrepancies, default=str)
    
    # Create combined content for hashing
    combined_content = data_json + discrepancies_json + timestamp
    
    # Generate SHA3-256 hash
    sha3_hash = hashlib.sha3_256(combined_content.encode()).hexdigest()
    
    # Save data and discrepancies to the snapshot directory
    data_path = os.path.join(snapshot_subdir, f"data_{timestamp}.parquet")
    discrepancies_path = os.path.join(snapshot_subdir, f"discrepancies_{timestamp}.json")
    hash_path = os.path.join(snapshot_subdir, f"hash_{timestamp}.txt")
    
    # Save data as parquet for efficient storage
    data.to_parquet(data_path, engine='pyarrow', index=False)
    
    # Save discrepancies as JSON
    with open(discrepancies_path, 'w') as f:
        json.dump(discrepancies, f, indent=2, default=str)
    
    # Save hash to file
    with open(hash_path, 'w') as f:
        f.write(f"SHA3-256: {sha3_hash}\n")
        f.write(f"Timestamp: {timestamp}\n")
        f.write(f"Records: {len(data)}\n")
        f.write(f"Discrepancies: {len(discrepancies)}\n")
    
    logger.info(f"Snapshot saved: {snapshot_subdir} with hash {sha3_hash[:16]}...")
    
    return snapshot_subdir

def verify_snapshot(snapshot_path: str, data: pd.DataFrame, 
                  discrepancies: List[Dict[str, Any]]) -> bool:
    """
    Verify that the data matches the snapshot hash
    """
    # Read the stored hash
    hash_files = [f for f in os.listdir(snapshot_path) if f.startswith('hash_')]
    if not hash_files:
        logger.error(f"No hash file found in {snapshot_path}")
        return False
    
    hash_file = os.path.join(snapshot_path, hash_files[0])
    with open(hash_file, 'r') as f:
        stored_hash = f.read().split(':')[1].split('\n')[0].strip()
    
    # Recalculate hash from provided data
    timestamp = os.path.basename(snapshot_path)
    data_json = data.to_json(orient='records', date_format='iso')
    discrepancies_json = json.dumps(discrepancies, default=str)
    combined_content = data_json + discrepancies_json + timestamp
    calculated_hash = hashlib.sha3_256(combined_content.encode()).hexdigest()
    
    return stored_hash == calculated_hash

def log_change(change_description: str, log_file: str = "data/archive/change_log.txt"):
    """
    Log a change to the audit trail
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {change_description}\n"
    
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    
    with open(log_file, 'a') as f:
        f.write(log_entry)
    
    logger.info(f"Change logged: {change_description}")

def get_available_snapshots(snapshot_dir: str = "data/archive/snapshots/") -> List[str]:
    """
    Get list of available snapshots sorted by date (most recent first)
    """
    if not os.path.exists(snapshot_dir):
        return []
    
    snapshots = [d for d in os.listdir(snapshot_dir) 
                 if os.path.isdir(os.path.join(snapshot_dir, d))]
    
    # Sort by date (assuming format YYYYMMDD_HHMMSS)
    snapshots.sort(reverse=True)
    return snapshots

def restore_from_snapshot(snapshot_path: str) -> tuple:
    """
    Restore data and discrepancies from a snapshot
    Returns (data: pd.DataFrame, discrepancies: List[Dict[str, Any]])
    """
    # Find data file
    data_files = [f for f in os.listdir(snapshot_path) if f.startswith('data_') and f.endswith('.parquet')]
    discrepancies_files = [f for f in os.listdir(snapshot_path) if f.startswith('discrepancies_') and f.endswith('.json')]
    
    if not data_files or not discrepancies_files:
        raise FileNotFoundError(f"Data or discrepancies file not found in {snapshot_path}")
    
    data_path = os.path.join(snapshot_path, data_files[0])
    discrepancies_path = os.path.join(snapshot_path, discrepancies_files[0])
    
    # Load data and discrepancies
    data = pd.read_parquet(data_path)
    
    with open(discrepancies_path, 'r') as f:
        discrepancies = json.load(f)
    
    logger.info(f"Restored from snapshot: {snapshot_path}")
    
    return data, discrepancies

class AuditTrail:
    """
    Class to manage the complete audit trail functionality
    """
    def __init__(self, base_dir: str = "data/archive/"):
        self.base_dir = base_dir
        self.snapshots_dir = os.path.join(base_dir, "snapshots/")
        self.log_file = os.path.join(base_dir, "change_log.txt")
        
        os.makedirs(self.snapshots_dir, exist_ok=True)
        os.makedirs(base_dir, exist_ok=True)
    
    def create_audit_point(self, data: pd.DataFrame, discrepancies: List[Dict[str, Any]], 
                          description: str = "") -> str:
        """
        Create an audit point (snapshot) with a description
        """
        snapshot_path = save_snapshot(data, discrepancies, self.base_dir)
        
        # Log the creation
        log_change(f"Created snapshot {os.path.basename(snapshot_path)}: {description}", self.log_file)
        
        return snapshot_path
    
    def list_audit_history(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        List recent audit history with timestamps and descriptions
        """
        snapshots = get_available_snapshots(self.snapshots_dir)
        history = []
        
        for snap in snapshots[:limit]:
            snap_path = os.path.join(self.snapshots_dir, snap)
            
            # Read hash file to get metadata
            hash_files = [f for f in os.listdir(snap_path) if f.startswith('hash_')]
            if hash_files:
                hash_file = os.path.join(snap_path, hash_files[0])
                with open(hash_file, 'r') as f:
                    lines = f.readlines()
                    hash_value = lines[0].split(':')[1].strip() if len(lines) > 0 else ""
                    record_count = lines[2].split(':')[1].strip() if len(lines) > 2 else "0"
                    disc_count = lines[3].split(':')[1].strip() if len(lines) > 3 else "0"
            
            history.append({
                'timestamp': snap,
                'hash_preview': hash_value[:16] if hash_value else "",
                'record_count': record_count,
                'discrepancy_count': disc_count,
                'path': snap_path
            })
        
        return history
    
    def get_change_log(self, limit: int = 20) -> List[str]:
        """
        Get recent entries from the change log
        """
        if not os.path.exists(self.log_file):
            return []
        
        with open(self.log_file, 'r') as f:
            lines = f.readlines()
        
        # Return last 'limit' lines
        return [line.strip() for line in lines[-limit:]]

def initialize_audit_system():
    """
    Initialize the audit system by creating necessary directories
    """
    base_dir = "data/archive/"
    os.makedirs(base_dir, exist_ok=True)
    os.makedirs(os.path.join(base_dir, "snapshots"), exist_ok=True)
    
    logger.info("Audit system initialized")