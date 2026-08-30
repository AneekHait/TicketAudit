"""
Configuration and constants for Sanity Check App.
"""
import os
import json
import pathlib
from typing import Dict, List, Any

# --- File Paths ---
# Absolute path relative to this file, so it resolves correctly regardless of CWD
CONFIG_FILE = str(pathlib.Path(__file__).parent / "config.json")

# --- Default Configuration ---
DEFAULT_CONFIG: Dict[str, Any] = {
    "null_threshold": 20,
    "desc_min_length": 20,
    "desc_max_length": 5000,
    # Native Excel PivotTables in the exported report. Assembled from raw OOXML,
    # so this is the escape hatch if a future Excel build rejects them: with it
    # off, every sheet is still a named Excel Table and Insert > PivotTable is a
    # three-click operation.
    "report_native_pivots": True,
    # Corrections to the automatic column matching, as {category: column}.
    # Not keyed by file, deliberately: an analyst works with one export shape, a
    # hash key would be unreadable in a file the in-app help tells people to edit
    # by hand, and an override naming a column the current file lacks is reported
    # as ignored rather than silently doing nothing.
    "column_overrides": {},
    # Pins how text dates are read: null infers, true is DD/MM, false is MM/DD.
    "date_dayfirst": None,
    # Last Extract selection, so extracting the same shape again is one click.
    # A plain list of column names: if a name is absent from the current file it
    # is simply not ticked, which is the right outcome for a different export.
    "extract_columns": [],
    "extract_date_column": None,
    "extract_format": "xlsx",
    "extract_month_mode": "range",
    "extract_exclude_cancelled": False,
    # --- UI state (remembered between sessions) ---
    "last_folder": "",
    # Must track gui.app_pyside.SIDEBAR_WIDTH. ConfigManager merges this dict
    # into whatever config.json holds, so the key is always present and the
    # `self.config.get("sidebar_width", SIDEBAR_WIDTH)` fallback in the GUI
    # never fires - this number, not the token, is what a fresh install gets.
    "sidebar_width": 244,
    # Stored as an explicit dict rather than a base64 saveGeometry() blob so
    # config.json stays hand-editable (the in-app LLM Tips dialog tells users
    # to edit it directly).
    "window_geometry": {},
    "last_view": ""
}

# --- Column Keywords for Validation (MANDATORY COLUMNS) ---
COLUMN_KEYWORDS: Dict[str, List[str]] = {
    "Ticket Identification": [
        "number", "ticket_id", "ticket_number", "incident_number", 
        "request_number", "case_id", "case_number", "sys_id",
        "inc_number", "sr_number", "cr_number", "ritm_number",
        "ticket_no", "incident_id", "request_id", "case_no",
        "reference", "ref_number", "ref_no", "ticket_ref"
    ],
    "Assignment Group": [
        "assignment_group", "assigned_group", "support_group", 
        "resolver_group", "assigned_to_group", "team",
        "group_name", "support_team", "resolver_team",
        "owning_group", "owner_group", "responsible_group"
    ],
    "Configuration Item": [
        "cmdb_ci", "ci_name", "configuration_item", 
        "affected_ci", "service_ci", "business_service",
        "service_name", "application", "app_name",
        "system", "component", "service_offering"
    ],
    "Priority": [
        "priority", "urgency", "severity", "impact",
        "priority_level", "sla_priority", "ticket_priority"
    ],
    "State/Status": [
        "state", "status", "incident_state", "ticket_state",
        "current_state", "workflow_state", "ticket_status",
        "request_state", "case_status", "resolution_status"
    ],
    "Created": [
        "opened_at", "created", "created_at", "created_date", 
        "open_date", "creation_date", "sys_created_on",
        "submitted_date", "reported_date", "raised_date",
        "logged_date", "opened_date", "start_date", "raised_on"
    ],
    "Closed": [
        "resolved_at", "closed_at", "closed", "resolved", 
        "close_date", "resolution_date", "closed_date",
        "completed_date", "completion_date", "end_date",
        "resolved_date", "closure_date", "closed_on"
    ],
    "Short Description": [
        "short_description", "title", "subject", "summary", 
        "brief", "headline", "short_desc", "issue_title",
        "ticket_title", "problem_summary"
    ],
    "Description": [
        "description", "long_description", "details", "notes", 
        "full_description", "issue_description", "problem_description",
        "ticket_description", "request_description", "comments"
    ],
}

# ID patterns for Ticket Identification (regex patterns)
ID_PATTERNS: List[str] = [
    r'^INC\d+',      # INC followed by any digits
    r'^CR\d+',       # CR followed by any digits
    r'^SR\d+',       # SR followed by any digits
    r'^RITM\d+',     # RITM followed by any digits
    r'^CHG\d+',      # CHG followed by any digits
    r'^PRB\d+',      # PRB followed by any digits
    r'^REQ\d+',      # REQ followed by any digits
    r'^TASK\d+',     # TASK followed by any digits
    r'^CTASK\d+',    # CTASK followed by any digits
    r'^SCTASK\d+',   # SCTASK followed by any digits
    r'^TKT\d+',      # TKT followed by any digits
    r'^CASE\d+',     # CASE followed by any digits
    r'^IM\d+',       # IM followed by any digits
    r'^SD\d+',       # SD followed by any digits
    r'^HDT\d+',      # HDT (Help Desk Ticket)
    r'^WO\d+',       # WO (Work Order)
    r'^\d+$',        # Pure numeric IDs (any length)
]

# Common value patterns for different column types
VALUE_PATTERNS: Dict[str, List[str]] = {
    "Priority": [
        r'^P[1-5]$',
        r'^[1-5]$',
        r'^(Critical|High|Medium|Low|Very Low)$',
        r'^(Urgent|Normal|Low)$',
    ],
    "State/Status": [
        r'^(New|Open|In Progress|Pending|Resolved|Closed|Cancelled)$',
        r'^(Active|Inactive|On Hold|Assigned|Work in Progress)$',
        r'^[1-8]$',
    ],
}

# Columns to EXCLUDE from certain matches (prevent false positives)
EXCLUDE_PATTERNS: Dict[str, List[str]] = {
    "Ticket Identification": [
        r'.*state.*', r'.*status.*', r'.*description.*', 
        r'.*group.*', r'.*date.*', r'.*priority.*'
    ],
    "Configuration Item": [
        r'.*state.*', r'.*status.*', r'.*incident.*(?!.*ci)',
        r'.*parent.*', r'.*child.*'
    ],
}

# Keywords for identifying ID/ticket number columns by VALUES
ID_KEYWORDS: List[str] = ["INC", "CR", "SR", "RITM", "CHG", "PRB", "REQ", "TASK"]

# Stopwords dropped when building the token signature used to group
# reworded descriptions of the same underlying issue.
DESCRIPTION_STOPWORDS: set = {
    "about", "after", "again", "also", "another", "back", "because", "been",
    "before", "being", "both", "cannot", "come", "could", "does", "doing",
    "done", "down", "during", "each", "even", "ever", "every", "from",
    "getting", "give", "goes", "going", "gone", "have", "having", "here",
    "into", "just", "know", "like", "made", "make", "many", "more", "most",
    "much", "must", "need", "needs", "only", "other", "over", "please",
    "same", "seems", "shall", "should", "since", "some", "still", "such",
    "than", "that", "their", "them", "then", "there", "these", "they",
    "thing", "this", "those", "through", "time", "trying", "under", "until",
    "upon", "used", "using", "very", "want", "well", "were", "what", "when",
    "where", "which", "while", "will", "with", "within", "without", "would",
    "your", "yours",
    # ITSM boilerplate that appears in nearly every ticket
    "team", "hello", "dear", "thanks", "thank", "regards", "kindly",
    "ticket", "incident", "request", "issue", "user", "customer",
}


class ConfigManager:
    """Handles loading and saving application configuration."""
    
    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.config = self.load()
    
    def load(self) -> Dict[str, Any]:
        """Load configuration from file, with defaults."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return {**DEFAULT_CONFIG, **json.load(f)}
            except (json.JSONDecodeError, IOError):
                return DEFAULT_CONFIG.copy()
        return DEFAULT_CONFIG.copy()
    
    def save(self) -> None:
        """Save current configuration to file."""
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=4)
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value."""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set a configuration value."""
        self.config[key] = value
    
    def update(self, updates: Dict[str, Any]) -> None:
        """Update multiple configuration values."""
        self.config.update(updates)
