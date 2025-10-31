"""
Email Communications API Client
Sends email data to Supabase database via API endpoint
"""

import requests
import json
from datetime import datetime
from typing import Dict, List, Optional, Union


class EmailAPIClient:
    """Client for sending email communications data to Supabase API"""
    
    def __init__(self, api_url: str, api_key: str, timeout: int = 30):
        """
        Initialize the API client
        
        Args:
            api_url: Full API endpoint URL (e.g., "https://your-domain.vercel.app/api/email-communications")
            api_key: Your API key for authentication
            timeout: Request timeout in seconds (default: 30)
        """
        self.api_url = api_url
        self.api_key = api_key
        self.timeout = timeout
        self.headers = {
            "Content-Type": "application/json",
            "X-API-Key": api_key
        }
    
    def send_single_email(self, email_data: Dict) -> Optional[Dict]:
        """
        Send a single email communication record
        
        Args:
            email_data: Dictionary containing email data with required fields:
                - project_name (str): Project identifier (e.g., "SHI", "HHI", "HMD")
                - vessel_name (str): Vessel/Hull name (e.g., "H3307")
                - sender_email (str): Sender's email address
                - recipient_email (str): Recipient's email address
                - email_type (str): "sent" or "received"
                
                Optional fields:
                - sender_company (str): Sender's company name
                - recipient_company (str): Recipient's company name
                - vendor_name (str): Vendor name
                - subject (str): Email subject line
                - sent_date (str): ISO 8601 format datetime
                - notes (str): Additional notes
        
        Returns:
            Dict containing the created record or None if failed
        """
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=email_data,
                timeout=self.timeout
            )
            
            if response.status_code == 201:
                result = response.json()
                print(f"✅ Success: {result['message']}")
                print(f"   Email ID: {result['data']['id']}")
                return result['data']
            elif response.status_code == 400:
                error = response.json()
                print(f"❌ Validation Error: {error['message']}")
                print(f"   Details: {error.get('error', 'N/A')}")
            elif response.status_code == 401:
                print("❌ Authentication failed. Check your API key.")
            elif response.status_code == 500:
                error = response.json()
                print(f"❌ Server Error: {error['message']}")
            else:
                print(f"❌ Unexpected error: {response.status_code}")
                print(f"   Response: {response.text}")
            
            return None
            
        except requests.exceptions.Timeout:
            print("❌ Request timed out")
            return None
        except requests.exceptions.ConnectionError:
            print("❌ Connection error. Check the API URL and your internet connection.")
            return None
        except Exception as e:
            print(f"❌ Unexpected error: {str(e)}")
            return None
    
    def send_bulk_emails(self, emails_list: List[Dict]) -> Optional[List[Dict]]:
        """
        Send multiple email communication records in one request
        
        Args:
            emails_list: List of email data dictionaries (same format as send_single_email)
        
        Returns:
            List of created records or None if failed
        """
        bulk_data = {"emails": emails_list}
        
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=bulk_data,
                timeout=self.timeout
            )
            
            if response.status_code == 201:
                result = response.json()
                print(f"✅ {result['message']}")
                print(f"   Saved {len(result['data'])} email(s)")
                return result['data']
            elif response.status_code == 400:
                error = response.json()
                print(f"❌ Validation Error: {error['message']}")
                print(f"   Details: {error.get('error', 'N/A')}")
            elif response.status_code == 401:
                print("❌ Authentication failed. Check your API key.")
            elif response.status_code == 500:
                error = response.json()
                print(f"❌ Server Error: {error['message']}")
            else:
                print(f"❌ Unexpected error: {response.status_code}")
                print(f"   Response: {response.text}")
            
            return None
            
        except requests.exceptions.Timeout:
            print("❌ Request timed out")
            return None
        except requests.exceptions.ConnectionError:
            print("❌ Connection error. Check the API URL and your internet connection.")
            return None
        except Exception as e:
            print(f"❌ Unexpected error: {str(e)}")
            return None


def load_config_from_file(config_file: str = "api_config.json") -> Dict:
    """
    Load API configuration from JSON file
    
    Args:
        config_file: Path to config file
    
    Returns:
        Dictionary with api_url and api_key
    """
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        print(f"❌ Config file '{config_file}' not found")
        return {}
    except json.JSONDecodeError:
        print(f"❌ Invalid JSON in config file '{config_file}'")
        return {}


# ============================================================================
# USAGE EXAMPLES
# ============================================================================

def example_single_email():
    """Example: Send a single email communication"""
    
    # Option 1: Direct configuration
    API_URL = "https://your-domain.vercel.app/api/email-communications"
    API_KEY = "your_api_key_here"
    
    # Option 2: Load from config file (uncomment to use)
    # config = load_config_from_file("api_config.json")
    # API_URL = config.get("api_url")
    # API_KEY = config.get("api_key")
    
    # Initialize client
    client = EmailAPIClient(API_URL, API_KEY)
    
    # Prepare email data
    email_data = {
        "project_name": "SHI",
        "vessel_name": "H3307",
        "sender_email": "vendor@company.com",
        "recipient_email": "manager@shi.kr",
        "email_type": "received",
        "sender_company": "Company Vendor",
        "recipient_company": "SHI",
        "vendor_name": "Company Vendor",
        "subject": "Equipment Update",
        "sent_date": datetime.now().isoformat() + "Z",
        "notes": "Follow-up required"
    }
    
    # Send email data
    result = client.send_single_email(email_data)
    
    if result:
        print(f"\n📋 Created Record:")
        print(f"   ID: {result['id']}")
        print(f"   Project: {result['project_name']}")
        print(f"   Vessel: {result['vessel_name']}")
        print(f"   Subject: {result.get('subject', 'N/A')}")


def example_bulk_emails():
    """Example: Send multiple email communications in bulk"""
    
    API_URL = "https://your-domain.vercel.app/api/email-communications"
    API_KEY = "your_api_key_here"
    
    # Initialize client
    client = EmailAPIClient(API_URL, API_KEY)
    
    # Prepare multiple emails
    emails = [
        {
            "project_name": "SHI",
            "vessel_name": "H3307",
            "sender_email": "vendor1@company.com",
            "recipient_email": "manager@shi.kr",
            "email_type": "received",
            "subject": "Equipment Specs",
            "sent_date": "2024-10-13T09:30:00Z"
        },
        {
            "project_name": "HHI",
            "vessel_name": "H2450",
            "sender_email": "vendor2@techco.com",
            "recipient_email": "team@hhi.kr",
            "email_type": "sent",
            "subject": "Installation Schedule",
            "sent_date": "2024-10-13T10:15:00Z"
        },
        {
            "project_name": "HMD",
            "vessel_name": "H3100",
            "sender_email": "supplier@parts.com",
            "recipient_email": "procurement@hmd.kr",
            "email_type": "received",
            "subject": "Parts Delivery Confirmation"
        }
    ]
    
    # Send bulk emails
    results = client.send_bulk_emails(emails)
    
    if results:
        print(f"\n📋 Created {len(results)} Records:")
        for idx, record in enumerate(results, 1):
            print(f"   {idx}. ID {record['id']}: {record['project_name']} - {record['vessel_name']}")


def example_minimal_email():
    """Example: Send email with only required fields"""
    
    API_URL = "https://your-domain.vercel.app/api/email-communications"
    API_KEY = "your_api_key_here"
    
    client = EmailAPIClient(API_URL, API_KEY)
    
    # Minimal required fields only
    email_data = {
        "project_name": "SHI",
        "vessel_name": "H3307",
        "sender_email": "test@example.com",
        "recipient_email": "admin@shi.kr",
        "email_type": "sent"
    }
    
    result = client.send_single_email(email_data)
    return result


if __name__ == "__main__":
    """
    Run this script to test the API integration
    
    Before running:
    1. Replace API_URL with your actual Vercel deployment URL
    2. Replace API_KEY with your actual API key
    3. Update email_data with your test data
    """
    
    print("=" * 70)
    print("Email Communications API Client - Test Script")
    print("=" * 70)
    print()
    
    # Uncomment the example you want to run:
    
    # Example 1: Send single email
    # example_single_email()
    
    # Example 2: Send bulk emails
    # example_bulk_emails()
    
    # Example 3: Send minimal email
    # example_minimal_email()
    
    print()
    print("=" * 70)
    print("ℹ️  To use this script:")
    print("   1. Update API_URL and API_KEY in the examples")
    print("   2. Uncomment one of the example functions above")
    print("   3. Run: python send_email_api.py")
    print("=" * 70)

