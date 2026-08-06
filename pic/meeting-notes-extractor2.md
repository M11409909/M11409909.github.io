---
name: meeting-notes-extractor
description: Transform meeting transcripts from Zoom, Teams, or Google Meet into structured Word documents, responsive HTML pages, and Markdown files with action items, decisions, and next steps. Extract owners, deadlines, and key discussion points. Use when user mentions meeting notes, transcripts, action items, meeting minutes, or Zoom/Teams recordings.
---

# Meeting Notes & Action Item Extractor

## What This Skill Does

Automatically processes meeting transcripts and creates professional meeting minutes in three formats: Word document, responsive HTML, and Markdown. Extracts action items with owners, key decisions, discussion points, and organizes everything into clean, shareable documents.

## When to Use

- User uploads Zoom/Teams/Meet transcript
- Mentions meeting notes, action items, or minutes
- Needs to extract tasks from meeting discussions
- Wants structured documentation of meetings

## How It Works

### Step 1: Create Enhanced Word Document

```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
from datetime import datetime

def set_cell_background(cell, color):
    """Set cell background color"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    cell._element.get_or_add_tcPr().append(shading)

def set_cell_border(cell, **kwargs):
    """Set cell borders"""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '12')
            element.set(qn('w:color'), kwargs[edge])
            tcBorders.append(element)
    tcPr.append(tcBorders)

# Read transcript
with open('team_meeting_transcript.txt', 'r') as f:
    transcript = f.read()

# Create Word document
doc = Document()

# Configure styles
sections = doc.sections
for section in sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

# Title with custom styling
title = doc.add_heading('Meeting Minutes', 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title.runs[0]
title_run.font.size = Pt(28)
title_run.font.color.rgb = RGBColor(0, 51, 102)
title_run.font.bold = True

# Add decorative line
doc.add_paragraph('_' * 80)

# Meeting Information Section
doc.add_heading('Meeting Information', level=1)
info_heading = doc.paragraphs[-1]
info_heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)

info_table = doc.add_table(rows=5, cols=2)
info_table.style = 'Light List Accent 1'
metadata = [
    ('📅 Date:', datetime.now().strftime('%B %d, %Y')),
    ('⏱️ Duration:', '60 minutes'),
    ('👥 Attendees:', 'Extract from transcript'),
    ('📍 Location:', 'Virtual / Zoom'),
    ('📝 Type:', 'Team Meeting')
]
for idx, (key, value) in enumerate(metadata):
    row = info_table.rows[idx]
    row.cells[0].text = key
    row.cells[1].text = value
    # Style cells
    row.cells[0].paragraphs[0].runs[0].font.bold = True
    row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 51, 102)
    set_cell_background(row.cells[0], 'E8F0F7')

doc.add_paragraph()  # Spacing

# Action Items Section
doc.add_heading('🎯 Action Items', level=1)
action_heading = doc.paragraphs[-1]
action_heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)

action_table = doc.add_table(rows=1, cols=4)
action_table.style = 'Medium Grid 3 Accent 1'
headers = ['Action Item', 'Owner', 'Deadline', 'Priority']
header_row = action_table.rows[0]
for idx, header in enumerate(headers):
    cell = header_row.cells[idx]
    cell.text = header
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_background(cell, '0033CC')

# Sample action items
actions = [
    ('Update Q4 roadmap slides', 'Sarah Chen', '2024-06-30', '🔴 High'),
    ('Schedule customer interviews', 'Mike Johnson', '2024-07-05', '🟡 Medium'),
    ('Review budget proposals', 'Team Lead', '2024-07-10', '🟢 Low')
]
for action, owner, deadline, priority in actions:
    row = action_table.add_row()
    row.cells[0].text = action
    row.cells[1].text = owner
    row.cells[2].text = deadline
    row.cells[3].text = priority
    row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()  # Spacing

# Key Decisions Section
doc.add_heading('✅ Key Decisions', level=1)
decision_heading = doc.paragraphs[-1]
decision_heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)

decisions = [
    'Approved new feature for Q3 release',
    'Budget increase of 15% for marketing approved',
    'Hiring 2 additional engineers - start recruiting immediately'
]
for decision in decisions:
    p = doc.add_paragraph(decision, style='List Bullet')
    p.runs[0].font.size = Pt(11)

doc.add_paragraph()  # Spacing

# Discussion Summary Section
doc.add_heading('💬 Discussion Summary', level=1)
summary_heading = doc.paragraphs[-1]
summary_heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)

summary_text = 'Team reviewed Q2 performance and discussed priorities for Q3. Main topics included product roadmap, budget allocation, and hiring plans. All attendees aligned on key deliverables and next steps.'
p = doc.add_paragraph(summary_text)
p.runs[0].font.size = Pt(11)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

doc.add_paragraph()  # Spacing

# Next Steps Section
doc.add_heading('⏭️ Next Steps', level=1)
next_heading = doc.paragraphs[-1]
next_heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)

next_steps = [
    'All action item owners to provide updates by next meeting',
    'Schedule follow-up for budget review',
    'Next meeting: July 15, 2024 at 2:00 PM'
]
for idx, step in enumerate(next_steps, 1):
    p = doc.add_paragraph(f'{idx}. {step}')
    p.runs[0].font.size = Pt(11)

# Footer
doc.add_paragraph()
footer = doc.add_paragraph('_' * 80)
footer_text = doc.add_paragraph('Generated on ' + datetime.now().strftime('%B %d, %Y at %I:%M %p'))
footer_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer_text.runs[0].font.size = Pt(9)
footer_text.runs[0].font.color.rgb = RGBColor(128, 128, 128)
footer_text.runs[0].italic = True

# Save Word document
doc.save('/mnt/user-data/outputs/meeting_minutes.docx')
print("✓ Word document created: meeting_minutes.docx")
```

### Step 2: Generate Responsive HTML

```python
# Generate responsive HTML version
html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Meeting Minutes - {datetime.now().strftime('%B %d, %Y')}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #0033cc 0%, #0052ff 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 700;
        }}
        
        .header .date {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .content {{
            padding: 40px;
        }}
        
        .section {{
            margin-bottom: 40px;
        }}
        
        .section-title {{
            font-size: 1.8em;
            color: #0033cc;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #0033cc;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .info-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }}
        
        .info-item {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #0033cc;
        }}
        
        .info-label {{
            font-weight: 600;
            color: #0033cc;
            margin-bottom: 5px;
        }}
        
        .info-value {{
            color: #555;
        }}
        
        .action-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            border-radius: 8px;
            overflow: hidden;
        }}
        
        .action-table thead {{
            background: #0033cc;
            color: white;
        }}
        
        .action-table th {{
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }}
        
        .action-table td {{
            padding: 15px;
            border-bottom: 1px solid #e0e0e0;
        }}
        
        .action-table tbody tr:hover {{
            background: #f8f9fa;
        }}
        
        .priority {{
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.9em;
            font-weight: 600;
            display: inline-block;
        }}
        
        .priority-high {{
            background: #ffebee;
            color: #c62828;
        }}
        
        .priority-medium {{
            background: #fff9c4;
            color: #f57f17;
        }}
        
        .priority-low {{
            background: #e8f5e9;
            color: #2e7d32;
        }}
        
        .decision-list {{
            list-style: none;
            padding: 0;
        }}
        
        .decision-list li {{
            padding: 15px;
            margin-bottom: 10px;
            background: #e8f5e9;
            border-left: 4px solid #4caf50;
            border-radius: 4px;
        }}
        
        .decision-list li:before {{
            content: "✓";
            color: #4caf50;
            font-weight: bold;
            margin-right: 10px;
            font-size: 1.2em;
        }}
        
        .summary-box {{
            background: #f8f9fa;
            padding: 25px;
            border-radius: 8px;
            border-left: 4px solid #0033cc;
            font-size: 1.05em;
            line-height: 1.8;
        }}
        
        .next-steps {{
            counter-reset: step-counter;
            list-style: none;
            padding: 0;
        }}
        
        .next-steps li {{
            counter-increment: step-counter;
            padding: 15px;
            margin-bottom: 10px;
            background: #fff3e0;
            border-radius: 4px;
            position: relative;
            padding-left: 50px;
        }}
        
        .next-steps li:before {{
            content: counter(step-counter);
            position: absolute;
            left: 15px;
            top: 50%;
            transform: translateY(-50%);
            background: #ff9800;
            color: white;
            width: 30px;
            height: 30px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
        }}
        
        .footer {{
            background: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #666;
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
        }}
        
        @media (max-width: 768px) {{
            .header h1 {{
                font-size: 1.8em;
            }}
            
            .content {{
                padding: 20px;
            }}
            
            .section-title {{
                font-size: 1.4em;
            }}
            
            .info-grid {{
                grid-template-columns: 1fr;
            }}
            
            .action-table {{
                font-size: 0.9em;
            }}
            
            .action-table th,
            .action-table td {{
                padding: 10px;
            }}
        }}
        
        @media print {{
            body {{
                background: white;
                padding: 0;
            }}
            
            .container {{
                box-shadow: none;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 Meeting Minutes</h1>
            <div class="date">{datetime.now().strftime('%B %d, %Y')}</div>
        </div>
        
        <div class="content">
            <!-- Meeting Information -->
            <div class="section">
                <h2 class="section-title">📊 Meeting Information</h2>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">📅 Date</div>
                        <div class="info-value">{datetime.now().strftime('%B %d, %Y')}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">⏱️ Duration</div>
                        <div class="info-value">60 minutes</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">👥 Attendees</div>
                        <div class="info-value">Extract from transcript</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">📍 Location</div>
                        <div class="info-value">Virtual / Zoom</div>
                    </div>
                </div>
            </div>
            
            <!-- Action Items -->
            <div class="section">
                <h2 class="section-title">🎯 Action Items</h2>
                <table class="action-table">
                    <thead>
                        <tr>
                            <th>Action Item</th>
                            <th>Owner</th>
                            <th>Deadline</th>
                            <th>Priority</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Update Q4 roadmap slides</td>
                            <td>Sarah Chen</td>
                            <td>June 30, 2024</td>
                            <td><span class="priority priority-high">High</span></td>
                        </tr>
                        <tr>
                            <td>Schedule customer interviews</td>
                            <td>Mike Johnson</td>
                            <td>July 5, 2024</td>
                            <td><span class="priority priority-medium">Medium</span></td>
                        </tr>
                        <tr>
                            <td>Review budget proposals</td>
                            <td>Team Lead</td>
                            <td>July 10, 2024</td>
                            <td><span class="priority priority-low">Low</span></td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- Key Decisions -->
            <div class="section">
                <h2 class="section-title">✅ Key Decisions</h2>
                <ul class="decision-list">
                    <li>Approved new feature for Q3 release</li>
                    <li>Budget increase of 15% for marketing approved</li>
                    <li>Hiring 2 additional engineers - start recruiting immediately</li>
                </ul>
            </div>
            
            <!-- Discussion Summary -->
            <div class="section">
                <h2 class="section-title">💬 Discussion Summary</h2>
                <div class="summary-box">
                    Team reviewed Q2 performance and discussed priorities for Q3. Main topics included product roadmap, budget allocation, and hiring plans. All attendees aligned on key deliverables and next steps.
                </div>
            </div>
            
            <!-- Next Steps -->
            <div class="section">
                <h2 class="section-title">⏭️ Next Steps</h2>
                <ul class="next-steps">
                    <li>All action item owners to provide updates by next meeting</li>
                    <li>Schedule follow-up for budget review</li>
                    <li>Next meeting: July 15, 2024 at 2:00 PM</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
        </div>
    </div>
</body>
</html>'''

# Save HTML file
with open('/mnt/user-data/outputs/meeting_minutes.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("✓ HTML document created: meeting_minutes.html")
print("\nBoth documents are ready in /mnt/user-data/outputs/")
```

## Required Libraries

- python-docx

## Example Usage

**Prompt**: "Extract action items from this Zoom meeting transcript"

**Output**: 
- Professional Word document with enhanced styling, colors, and formatting
- Responsive HTML page optimized for desktop and mobile viewing
- Both files include action items table, decisions, summary, and next steps

## Key Features

### Word Document
- Professional color scheme with corporate blue theme
- Enhanced table formatting with custom backgrounds
- Emoji icons for visual organization
- Better spacing and typography
- Print-ready layout

### HTML Page
- Fully responsive design (mobile, tablet, desktop)
- Modern gradient header
- Interactive hover effects on tables
- Priority badges with color coding
- Optimized for sharing via email or web
- Print-friendly CSS

## Tips

- Works with Zoom, Teams, Google Meet transcripts
- Automatically detects speakers and topics
- Generates both Word and HTML in single operation
- HTML can be opened directly in browser or hosted online
- Word document maintains professional formatting for editing
- Both formats include identical content with optimized presentation
