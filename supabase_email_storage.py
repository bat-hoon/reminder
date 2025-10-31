"""
Supabase Email Storage Module
Stores email communications to Supabase database for emails matching yard tag pattern.
This module is designed to be imported and called from Auto_Reminder_List.py without modifying base code.
"""

import re
import json
import os
from datetime import datetime
from typing import Optional, Dict, List
from zoneinfo import ZoneInfo

try:
    from supabase import create_client, Client
except ImportError:
    print("[WARN] supabase-py library not installed. Run: pip install supabase")
    Client = None


# Yard tag pattern matching parse_yard_tag() function from Auto_Reminder_List.py
YARD_TAG_PATTERN = re.compile(r"\[(SHI|HMD|HHI|HSHI|HO|HJSC)(\d+)(MIN|H|D|W|M)\]", re.IGNORECASE)

# Vessel name patterns (alphanumeric codes like SN0000, H3307, H2450, etc.)
VESSEL_NAME_PATTERNS = [
    re.compile(r'\b([A-Z]{1,3}\d{3,5})\b', re.IGNORECASE),  # SN2693, H3307, H2450
    re.compile(r'\b([A-Z]\.\d{3,5})\b', re.IGNORECASE),     # H.2378
    re.compile(r'\b([A-Z]\d{4,6})\b', re.IGNORECASE),       # H3100
]


def load_supabase_config(config_file: str = "supabase_config.json") -> Dict:
    """
    Load Supabase configuration from JSON file.
    
    Args:
        config_file: Path to config file
        
    Returns:
        Dictionary with supabase_url and supabase_service_key
    """
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        print(f"[WARN] Supabase config file '{config_file}' not found")
        return {}
    except json.JSONDecodeError:
        print(f"[ERR] Invalid JSON in config file '{config_file}'")
        return {}


def validate_yard_tag_pattern(subject: str) -> bool:
    """
    Validate that subject contains yard tag pattern.
    
    Args:
        subject: Email subject line
        
    Returns:
        True if subject contains valid yard tag pattern, False otherwise
    """
    if not subject:
        return False
    subject_upper = subject.upper().replace("［", "[").replace("］", "]")
    return bool(YARD_TAG_PATTERN.search(subject_upper))


def extract_vessel_name(subject: str) -> Optional[str]:
    """
    Extract vessel name from subject line.
    Looks for patterns like SN0000, H3307, H.2378, etc.
    
    Args:
        subject: Email subject line
        
    Returns:
        Vessel name if found, None otherwise
    """
    if not subject:
        return None
    
    # Try each pattern
    for pattern in VESSEL_NAME_PATTERNS:
        match = pattern.search(subject)
        if match:
            vessel = match.group(1)
            # Clean up common prefixes/suffixes
            vessel = vessel.strip().upper()
            return vessel
    
    return None


def to_local_naive(dt) -> Optional[datetime]:
    """
    Convert Outlook datetime to local naive datetime.
    Compatible with Auto_Reminder_List.py function.
    """
    if not dt:
        return None
    try:
        if isinstance(dt, datetime):
            if dt.tzinfo is None:
                return dt
            # Convert to local timezone naive
            local_tz = ZoneInfo("Asia/Seoul")  # Adjust if needed
            return dt.astimezone(local_tz).replace(tzinfo=None)
        return None
    except Exception:
        return None


def extract_recipient_emails(mail_item) -> List[str]:
    """
    Extract recipient email addresses from mail item.
    
    Args:
        mail_item: Outlook mail item
        
    Returns:
        List of recipient email addresses (Type 1 or 3)
    """
    recipients = []
    try:
        recips = getattr(mail_item, "Recipients", None)
        if not recips:
            return recipients
        
        for i in range(1, recips.Count + 1):
            try:
                r = recips.Item(i)
                if getattr(r, "Type", 1) in (1, 3):  # To (1) or BCC (3)
                    addr = getattr(r, "Address", None) or getattr(r, "Name", None)
                    if addr:
                        recipients.append(str(addr).lower().strip())
            except Exception:
                continue
    except Exception:
        pass
    
    return recipients


def get_sender_email(mail_item) -> Optional[str]:
    """
    Extract sender email address from mail item.
    
    Args:
        mail_item: Outlook mail item
        
    Returns:
        Sender email address or None
    """
    try:
        addr = getattr(mail_item, "SenderEmailAddress", None)
        if addr:
            return str(addr).lower().strip()
    except Exception:
        pass
    
    try:
        sender = getattr(mail_item, "Sender", None)
        if sender:
            addr = getattr(sender, "Address", None) or getattr(sender, "Name", None)
            if addr:
                return str(addr).lower().strip()
    except Exception:
        pass
    
    return None


def store_email_to_supabase(mail_item, yard_code: str, verbose: bool = False) -> bool:
    """
    Store email communication to Supabase database.
    
    This function extracts email data and stores it to Supabase, but only if:
    1. Subject contains valid yard tag pattern
    2. Supabase is configured properly
    
    Args:
        mail_item: Outlook mail item
        yard_code: Yard code extracted from subject (e.g., "SHI", "HMD", "HHI")
        verbose: Whether to print verbose logging
        
    Returns:
        True if successfully stored, False otherwise
    """
    if Client is None:
        if verbose:
            print("[WARN] Supabase client not available. Install: pip install supabase")
        return False
    
    # Load configuration
    config = load_supabase_config()
    supabase_url = config.get("supabase_url")
    supabase_key = config.get("supabase_service_key")
    
    if not supabase_url or not supabase_key:
        if verbose:
            print("[WARN] Supabase not configured. Create supabase_config.json")
        return False
    
    try:
        # Initialize Supabase client
        supabase: Client = create_client(supabase_url, supabase_key)
    except Exception as e:
        if verbose:
            print(f"[ERR] Failed to create Supabase client: {e}")
        return False
    
    try:
        # Extract basic email data
        subject = getattr(mail_item, "Subject", None) or ""
        sent_date = to_local_naive(getattr(mail_item, "SentOn", None))
        
        # Validate yard tag pattern
        if not validate_yard_tag_pattern(subject):
            if verbose:
                print(f"[SKIP] Subject does not contain yard tag pattern: {subject[:50]}")
            return False
        
        # Extract sender email
        sender_email = get_sender_email(mail_item)
        if not sender_email:
            if verbose:
                print("[SKIP] Could not extract sender email")
            return False
        
        # Extract recipients
        recipient_emails = extract_recipient_emails(mail_item)
        if not recipient_emails:
            if verbose:
                print("[SKIP] No recipients found")
            return False
        
        # Extract vessel name from subject
        vessel_name = extract_vessel_name(subject)
        if not vessel_name:
            # Fallback: use first part of subject or "UNKNOWN"
            vessel_name = subject.split()[0] if subject.split() else "UNKNOWN"
        
        # Use yard_code as project_name
        project_name = yard_code.upper()
        
        # Convert sent_date to ISO format with timezone
        if sent_date:
            sent_date_iso = sent_date.isoformat() + "Z"
        else:
            sent_date_iso = datetime.now().isoformat() + "Z"
        
        # Store each recipient as a separate record
        success_count = 0
        for recipient_email in recipient_emails:
            try:
                email_data = {
                    "project_name": project_name,
                    "vessel_name": vessel_name,
                    "sender_email": sender_email,
                    "recipient_email": recipient_email,
                    "subject": subject,
                    "sent_date": sent_date_iso,
                    "email_type": "sent"
                }
                
                # Insert into Supabase
                result = supabase.table("email_communications").insert(email_data).execute()
                
                if result.data:
                    success_count += 1
                    if verbose:
                        print(f"[DB] Stored email: {project_name}/{vessel_name} | {sender_email} → {recipient_email}")
                else:
                    if verbose:
                        print(f"[WARN] Failed to store email for recipient: {recipient_email}")
                        
            except Exception as e:
                if verbose:
                    print(f"[ERR] Error storing email for {recipient_email}: {e}")
                continue
        
        if success_count > 0:
            if verbose:
                print(f"[SUCCESS] Stored {success_count} email record(s) to Supabase")
            return True
        else:
            return False
            
    except Exception as e:
        if verbose:
            print(f"[ERR] Error in store_email_to_supabase: {e}")
        return False

