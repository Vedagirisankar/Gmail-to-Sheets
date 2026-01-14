"""
Email Parser Module
Extracts relevant information from Gmail messages
"""

import base64
import re
from datetime import datetime
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser


class HTMLToTextParser(HTMLParser):
    """Convert HTML to plain text"""
    
    def __init__(self):
        super().__init__()
        self.text = []
    
    def handle_data(self, data):
        self.text.append(data)
    
    def get_text(self):
        return ' '.join(self.text).strip()


class EmailParser:
    """Parse Gmail API email messages"""
    
    def parse_email(self, email):
        """
        Parse email message from Gmail API
        
        Args:
            email: Email message object from Gmail API
            
        Returns:
            Dictionary with from, subject, date, content
        """
        try:
            headers = email['payload']['headers']
            
            # Extract headers
            from_addr = self._get_header(headers, 'From')
            subject = self._get_header(headers, 'Subject')
            date = self._get_header(headers, 'Date')
            
            # Parse and format date
            formatted_date = self._format_date(date)
            
            # Extract email body
            content = self._extract_body(email['payload'])
            
            return {
                'from': from_addr,
                'subject': subject,
                'date': formatted_date,
                'content': content
            }
            
        except Exception as e:
            print(f"Error parsing email: {str(e)}")
            return None
    
    def _get_header(self, headers, name):
        """Extract specific header value"""
        for header in headers:
            if header['name'].lower() == name.lower():
                return header['value']
        return ''
    
    def _format_date(self, date_str):
        """
        Format date string to readable format
        
        Args:
            date_str: Date string from email header
            
        Returns:
            Formatted date string
        """
        try:
            dt = parsedate_to_datetime(date_str)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return date_str
    
    def _extract_body(self, payload):
        """
        Extract email body from payload
        Handles both plain text and HTML
        
        Args:
            payload: Email payload from Gmail API
            
        Returns:
            Plain text body
        """
        body = ''
        
        # Check if payload has parts (multipart email)
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    body = self._decode_body(part['body'])
                    break
                elif part['mimeType'] == 'text/html' and not body:
                    html_body = self._decode_body(part['body'])
                    body = self._html_to_text(html_body)
        else:
            # Single part email
            if 'body' in payload and 'data' in payload['body']:
                body = self._decode_body(payload['body'])
                
                # Convert HTML to text if needed
                if payload.get('mimeType') == 'text/html':
                    body = self._html_to_text(body)
        
        # Clean up the body
        body = self._clean_body(body)
        
        return body
    
    def _decode_body(self, body):
        """Decode base64 encoded body"""
        if 'data' in body:
            try:
                decoded = base64.urlsafe_b64decode(body['data']).decode('utf-8')
                return decoded
            except:
                return ''
        return ''
    
    def _html_to_text(self, html):
        """Convert HTML to plain text"""
        try:
            parser = HTMLToTextParser()
            parser.feed(html)
            return parser.get_text()
        except:
            # Fallback: simple tag removal
            return re.sub('<[^<]+?>', '', html)
    
    def _clean_body(self, body):
        """Clean up email body text"""
        # Remove excessive whitespace
        body = re.sub(r'\n\s*\n', '\n\n', body)
        body = re.sub(r' +', ' ', body)
        
        # Limit length to avoid sheet cell limits
        max_length = 5000
        if len(body) > max_length:
            body = body[:max_length] + '... [truncated]'
        
        return body.strip()