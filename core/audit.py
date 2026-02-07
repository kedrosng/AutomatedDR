"""
Audit Trail Module for JV Cost Reconciliation

Implements:
- SHA3-256 snapshot creation with date-based archival
- Change logging with timestamps and user actions
- Snapshot restoration functionality
"""

import json
import logging
import hashlib
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field, asdict
from datetime import datetime
import shutil

import pandas as pd
import yaml

# Configure logging
logger = logging.getLogger(__name__)


@dataclass
class SnapshotMetadata:
    """Metadata for a data snapshot."""
    snapshot_id: str
    timestamp: str
    hash_sha3: str
    file_count: int
    total_rows: int
    partner: str
    source_type: str
    description: str
    created_by: str = "system"


@dataclass
class ChangeLogEntry:
    """Entry for the change log."""
    timestamp: str
    action: str
    entity_type: str
    entity_id: str
    description: str
    user: str = "system"
    details: Dict[str, Any] = field(default_factory=dict)


class AuditTrail:
    """
    Manages audit trail with snapshots and change logging.
    
    Supports:
    - SHA3-256 hashed snapshots saved to archive
    - JSON or SQLite change logging
    - Snapshot restoration
    """
    
    def __init__(
        self, 
        archive_path: Path,
        rules_path: Optional[Path] = None,
        use_sqlite: bool = False
    ):
        """
        Initialize the audit trail manager.
        
        Args:
            archive_path: Directory for storing snapshots
            rules_path: Path to rules.yaml for configuration
            use_sqlite: Use SQLite for change log (vs JSON)
        """
        self.archive_path = Path(archive_path)
        self.archive_path.mkdir(parents=True, exist_ok=True)
        
        self.use_sqlite = use_sqlite
        
        # Load rules if provided
        if rules_path:
            with open(rules_path, 'r', encoding='utf-8') as f:
                rules = yaml.safe_load(f)
                audit_config = rules.get('audit', {})
                self.retention_days = audit_config.get('snapshot_retention_days', 90)
                self.use_sqlite = audit_config.get('log_format', 'json') == 'sqlite'
        else:
            self.retention_days = 90
            
        # Initialize logging backend
        self._init_logging_backend()
        
    def _init_logging_backend(self) -> None:
        """Initialize the change log storage backend."""
        if self.use_sqlite:
            self.db_path = self.archive_path / "audit_log.db"
            self._init_sqlite()
        else:
            self.log_path = self.archive_path / "change_log.json"
            self._change_log: List[Dict] = []
            self._load_json_log()
            
    def _init_sqlite(self) -> None:
        """Initialize SQLite database for audit logging."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS change_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                action TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                entity_id TEXT NOT NULL,
                description TEXT,
                user TEXT DEFAULT 'system',
                details TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS snapshots (
                snapshot_id TEXT PRIMARY KEY,
                timestamp TEXT NOT NULL,
                hash_sha3 TEXT NOT NULL,
                file_count INTEGER,
                total_rows INTEGER,
                partner TEXT,
                source_type TEXT,
                description TEXT,
                created_by TEXT DEFAULT 'system'
            )
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_timestamp ON change_log(timestamp)
        ''')
        
        conn.commit()
        conn.close()
        logger.info(f"SQLite audit log initialized at {self.db_path}")
        
    def _load_json_log(self) -> None:
        """Load existing JSON change log."""
        if self.log_path.exists():
            try:
                with open(self.log_path, 'r', encoding='utf-8') as f:
                    self._change_log = json.load(f)
                logger.info(f"Loaded {len(self._change_log)} log entries")
            except Exception as e:
                logger.warning(f"Could not load change log: {e}")
                self._change_log = []
        else:
            self._change_log = []
            
    def _save_json_log(self) -> None:
        """Save change log to JSON file."""
        try:
            with open(self.log_path, 'w', encoding='utf-8') as f:
                json.dump(self._change_log, f, indent=2, default=str)
        except Exception as e:
            logger.error(f"Failed to save change log: {e}")
            
    def _compute_hash(self, data: bytes) -> str:
        """
        Compute SHA3-256 hash of data.
        
        Args:
            data: Bytes to hash
            
        Returns:
            Hex-encoded hash string
        """
        hasher = hashlib.sha3_256()
        hasher.update(data)
        return hasher.hexdigest()
    
    def create_snapshot(
        self,
        df: pd.DataFrame,
        partner: str,
        source_type: str,
        description: str = "",
        user: str = "system"
    ) -> SnapshotMetadata:
        """
        Create a snapshot of DataFrame with SHA3-256 hash.
        
        Args:
            df: DataFrame to snapshot
            partner: Partner identifier
            source_type: Type of data source
            description: Human-readable description
            user: User creating the snapshot
            
        Returns:
            SnapshotMetadata for the created snapshot
        """
        timestamp = datetime.now()
        date_str = timestamp.strftime("%Y%m%d")
        time_str = timestamp.strftime("%H%M%S")
        
        # Create snapshot directory
        snapshot_dir = self.archive_path / "snapshots" / date_str
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate snapshot ID
        snapshot_id = f"{partner}_{source_type}_{date_str}_{time_str}"
        
        # Serialize DataFrame
        csv_bytes = df.to_csv(index=False).encode('utf-8')
        hash_value = self._compute_hash(csv_bytes)
        
        # Save snapshot file
        snapshot_file = snapshot_dir / f"{snapshot_id}.csv"
        with open(snapshot_file, 'wb') as f:
            f.write(csv_bytes)
            
        # Create metadata
        metadata = SnapshotMetadata(
            snapshot_id=snapshot_id,
            timestamp=timestamp.isoformat(),
            hash_sha3=hash_value,
            file_count=1,
            total_rows=len(df),
            partner=partner,
            source_type=source_type,
            description=description,
            created_by=user
        )
        
        # Save metadata
        metadata_file = snapshot_dir / f"{snapshot_id}_meta.json"
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(asdict(metadata), f, indent=2)
            
        # Log the action
        self.log_change(
            action="create_snapshot",
            entity_type="snapshot",
            entity_id=snapshot_id,
            description=f"Created snapshot: {description or 'No description'}",
            user=user,
            details={"hash": hash_value, "rows": len(df)}
        )
        
        logger.info(f"Created snapshot {snapshot_id} with hash {hash_value[:16]}...")
        return metadata
    
    def log_change(
        self,
        action: str,
        entity_type: str,
        entity_id: str,
        description: str,
        user: str = "system",
        details: Optional[Dict[str, Any]] = None
    ) -> None:
        """
        Log a change action.
        
        Args:
            action: Type of action (create, update, resolve, etc.)
            entity_type: Type of entity affected
            entity_id: ID of affected entity
            description: Human-readable description
            user: User performing the action
            details: Additional details as dict
        """
        entry = ChangeLogEntry(
            timestamp=datetime.now().isoformat(),
            action=action,
            entity_type=entity_type,
            entity_id=entity_id,
            description=description,
            user=user,
            details=details or {}
        )
        
        if self.use_sqlite:
            self._log_to_sqlite(entry)
        else:
            self._change_log.append(asdict(entry))
            self._save_json_log()
            
    def _log_to_sqlite(self, entry: ChangeLogEntry) -> None:
        """Log change entry to SQLite."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO change_log (timestamp, action, entity_type, entity_id, description, user, details)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            entry.timestamp,
            entry.action,
            entry.entity_type,
            entry.entity_id,
            entry.description,
            entry.user,
            json.dumps(entry.details)
        ))
        
        conn.commit()
        conn.close()
        
    def get_change_log(
        self,
        limit: int = 100,
        entity_type: Optional[str] = None,
        action: Optional[str] = None
    ) -> List[ChangeLogEntry]:
        """
        Retrieve change log entries.
        
        Args:
            limit: Maximum entries to return
            entity_type: Filter by entity type
            action: Filter by action type
            
        Returns:
            List of ChangeLogEntry objects
        """
        if self.use_sqlite:
            return self._get_log_from_sqlite(limit, entity_type, action)
        else:
            entries = self._change_log.copy()
            
            if entity_type:
                entries = [e for e in entries if e.get('entity_type') == entity_type]
            if action:
                entries = [e for e in entries if e.get('action') == action]
                
            # Sort by timestamp descending
            entries.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
            
            return [
                ChangeLogEntry(**e) for e in entries[:limit]
            ]
            
    def _get_log_from_sqlite(
        self,
        limit: int,
        entity_type: Optional[str],
        action: Optional[str]
    ) -> List[ChangeLogEntry]:
        """Get log entries from SQLite."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        query = "SELECT timestamp, action, entity_type, entity_id, description, user, details FROM change_log WHERE 1=1"
        params = []
        
        if entity_type:
            query += " AND entity_type = ?"
            params.append(entity_type)
        if action:
            query += " AND action = ?"
            params.append(action)
            
        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(limit)
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        entries = []
        for row in rows:
            entries.append(ChangeLogEntry(
                timestamp=row[0],
                action=row[1],
                entity_type=row[2],
                entity_id=row[3],
                description=row[4],
                user=row[5],
                details=json.loads(row[6]) if row[6] else {}
            ))
            
        return entries
    
    def list_snapshots(
        self,
        partner: Optional[str] = None,
        source_type: Optional[str] = None,
        limit: int = 50
    ) -> List[SnapshotMetadata]:
        """
        List available snapshots.
        
        Args:
            partner: Filter by partner
            source_type: Filter by source type
            limit: Maximum snapshots to return
            
        Returns:
            List of SnapshotMetadata objects
        """
        snapshots = []
        snapshot_base = self.archive_path / "snapshots"
        
        if not snapshot_base.exists():
            return snapshots
            
        # Iterate through date directories
        for date_dir in sorted(snapshot_base.iterdir(), reverse=True):
            if not date_dir.is_dir():
                continue
                
            for meta_file in date_dir.glob("*_meta.json"):
                try:
                    with open(meta_file, 'r', encoding='utf-8') as f:
                        meta_dict = json.load(f)
                        meta = SnapshotMetadata(**meta_dict)
                        
                        # Apply filters
                        if partner and meta.partner != partner:
                            continue
                        if source_type and meta.source_type != source_type:
                            continue
                            
                        snapshots.append(meta)
                        
                        if len(snapshots) >= limit:
                            return snapshots
                            
                except Exception as e:
                    logger.warning(f"Could not load snapshot metadata {meta_file}: {e}")
                    
        return snapshots
    
    def restore_snapshot(
        self,
        snapshot_id: str
    ) -> Optional[pd.DataFrame]:
        """
        Restore a snapshot by ID.
        
        Args:
            snapshot_id: ID of snapshot to restore
            
        Returns:
            Restored DataFrame or None if not found
        """
        # Find the snapshot file
        snapshot_base = self.archive_path / "snapshots"
        
        for date_dir in snapshot_base.iterdir():
            if not date_dir.is_dir():
                continue
                
            snapshot_file = date_dir / f"{snapshot_id}.csv"
            meta_file = date_dir / f"{snapshot_id}_meta.json"
            
            if snapshot_file.exists():
                try:
                    # Load and verify hash
                    with open(snapshot_file, 'rb') as f:
                        content = f.read()
                        
                    current_hash = self._compute_hash(content)
                    
                    # Check stored hash
                    if meta_file.exists():
                        with open(meta_file, 'r', encoding='utf-8') as f:
                            meta = json.load(f)
                            stored_hash = meta.get('hash_sha3', '')
                            
                            if current_hash != stored_hash:
                                logger.warning(
                                    f"Hash mismatch for snapshot {snapshot_id}. "
                                    f"File may have been modified!"
                                )
                                
                    # Load DataFrame
                    df = pd.read_csv(snapshot_file)
                    
                    # Log the restore action
                    self.log_change(
                        action="restore_snapshot",
                        entity_type="snapshot",
                        entity_id=snapshot_id,
                        description=f"Restored snapshot {snapshot_id}",
                        details={"rows": len(df)}
                    )
                    
                    logger.info(f"Restored snapshot {snapshot_id} with {len(df)} rows")
                    return df
                    
                except Exception as e:
                    logger.error(f"Failed to restore snapshot {snapshot_id}: {e}")
                    return None
                    
        logger.warning(f"Snapshot {snapshot_id} not found")
        return None
    
    def cleanup_old_snapshots(self) -> int:
        """
        Remove snapshots older than retention period.
        
        Returns:
            Number of snapshots removed
        """
        from datetime import timedelta
        
        cutoff = datetime.now() - timedelta(days=self.retention_days)
        cutoff_str = cutoff.strftime("%Y%m%d")
        
        removed = 0
        snapshot_base = self.archive_path / "snapshots"
        
        if not snapshot_base.exists():
            return 0
            
        for date_dir in snapshot_base.iterdir():
            if not date_dir.is_dir():
                continue
                
            if date_dir.name < cutoff_str:
                try:
                    shutil.rmtree(date_dir)
                    removed += 1
                    logger.info(f"Removed old snapshot directory: {date_dir.name}")
                except Exception as e:
                    logger.error(f"Failed to remove {date_dir}: {e}")
                    
        if removed > 0:
            self.log_change(
                action="cleanup",
                entity_type="snapshot",
                entity_id="batch",
                description=f"Removed {removed} old snapshot directories",
                details={"cutoff_date": cutoff_str}
            )
            
        return removed
