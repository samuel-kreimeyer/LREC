#!/usr/bin/env python3
"""
Script to generate notification messages from spreadsheet data using Jinja template.
"""

import pandas as pd
import argparse
from jinja2 import Template
from pathlib import Path
from datetime import datetime


def main():
    parser = argparse.ArgumentParser(description='Generate notice messages from spreadsheet')
    parser.add_argument('spreadsheet', help='Path to spreadsheet file (CSV or Excel)')
    parser.add_argument('--bio', default='', help='Speaker bio (optional)')
    parser.add_argument('--lunch-provided', action='store_true', 
                       help='Use "Lunch will be provided." instead of default message')
    parser.add_argument('--output', '-o', default='notices.txt', 
                       help='Output file path (default: notices.txt)')
    parser.add_argument('--template', default='notice_template', 
                       help='Template file path (default: notice_template)')
    
    args = parser.parse_args()
    
    # Set lunch message
    lunch_message = "Lunch will be provided." if args.lunch_provided else "Feel free to bring your own lunch."
    
    # Read spreadsheet
    file_ext = Path(args.spreadsheet).suffix.lower()
    if file_ext in ['.xlsx', '.xls']:
        df = pd.read_excel(args.spreadsheet)
    else:
        df = pd.read_csv(args.spreadsheet)
    
    # Convert date column to datetime
    df['date'] = pd.to_datetime(df['date'])
    
    # Get current date
    current_date = datetime.now()
    
    # Filter for future dates only
    future_events = df[df['date'] > current_date]
    
    if future_events.empty:
        print("No future events found in the spreadsheet.")
        return
    
    # Find the event with the closest future date
    closest_event = future_events.loc[future_events['date'].idxmin()]
    
    # Load template
    with open(args.template, 'r') as f:
        template_content = f.read()
    
    template = Template(template_content)
    
    # Generate notice for the closest future event
    notice = template.render(
        date=closest_event['date'].strftime('%Y-%m-%d'),
        topic=closest_event['topic'],
        speaker=closest_event['speaker'],
        location=closest_event['location'],
        time=closest_event['time'],
        bio=args.bio,
        lunch_message=lunch_message
    )
    
    # Write output
    with open(args.output, 'w') as f:
        f.write(notice)
    
    print(f"Generated notice for {closest_event['date'].strftime('%Y-%m-%d')} event and saved to {args.output}")


if __name__ == '__main__':
    main()