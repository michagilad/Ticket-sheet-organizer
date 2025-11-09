#!/usr/bin/env python3
"""
Ticket Organizer Web Application
A Flask web app with a UI for uploading CSV files and downloading organized XLSX output.
"""

import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session, jsonify, after_this_request
from werkzeug.utils import secure_filename
import csv
import json
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-change-in-production-please-use-random-string')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'csv'}

# Available assignees list
ASSIGNEES = [
    'Cs. Zsófi',
    'Ricsi',
    'Attila',
    'Lili',
    'Flóra',
    'Csabi',
    'Marci',
    'Gábor',
    'Bálint',
    'B.Zsófi',
    'Muki',
    'Péter',
    'Adri',
    'Era',
    'Endre',
    'Vencel',
    'Áron'
]


def allowed_file(filename):
    """Check if file has allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def read_csv_robust(file_path):
    """Read CSV file with robust handling of line breaks and special characters."""
    with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    return rows


def extract_and_reorder_columns(rows):
    """Extract only the required columns and reorder them as specified."""
    # Define the columns we need from input
    input_columns = [
        'Ticket ID',
        'Backstage Experience page',
        'Company',
        'Create date',
        'Priority',
        'Record source detail 1',
        'Ticket description',
        'Ticket name',
        'Ticket status',
        'Associated Experience IDs',
        'Associated Experience',
        'Assignee',
        'Invalid ticket',
        'Ticket Category',
        'Sub Category',
        'Non-ticketed issues',
        'Additional Notes',
        'Notes - Admin',
        'Next step',
        'Last Updated - Status',
        'Last Updated - Next step',
        'Last Updated - Invalid Ticket',
        'Last Updated - Non Ticketed Issues',
        'Public Preview Link'  # This will be generated, not from input
    ]
    
    # Define the output order
    output_order = [
        'Associated Experience',
        'Backstage Experience page',
        'Public Preview Link',
        'Associated Experience IDs',
        'Assignee',
        'Ticket name',
        'Ticket description',
        'Ticket status',
        'Invalid ticket',
        'Ticket Category',
        'Sub Category',
        'Non-ticketed issues',
        'Additional Notes',
        'Notes - Admin',
        'Next step',
        'Create date',
        'Ticket ID',
        'Record source detail 1',
        'Company',
        'Priority',
        'Last Updated - Status',
        'Last Updated - Next step',
        'Last Updated - Invalid Ticket',
        'Last Updated - Non Ticketed Issues'
    ]
    
    processed_rows = []
    for row in rows:
        processed_row = {}
        for col in output_order:
            processed_row[col] = row.get(col, '')
        
        # If Associated Experience is empty, use Experience Name as fallback
        if not processed_row.get('Associated Experience') or not str(processed_row['Associated Experience']).strip():
            experience_name = row.get('Experience Name', '')
            if experience_name:
                processed_row['Associated Experience'] = experience_name
        
        processed_rows.append(processed_row)
    
    # Sort by Associated Experience IDs (numerically) before grouping
    # This ensures that all rows with the same Experience ID are consecutive
    def get_experience_id(row):
        exp_id = row.get('Associated Experience IDs', '')
        if not exp_id or not str(exp_id).strip():
            # Empty IDs go to the end
            return float('inf')
        try:
            # Try to convert to integer for proper numeric sorting
            return int(exp_id)
        except (ValueError, TypeError):
            # If conversion fails, put at the end
            return float('inf')
    
    processed_rows.sort(key=get_experience_id)
    
    return processed_rows, output_order


def assign_experiences_to_assignees(groups, assignee_distribution):
    """
    Assign experiences to assignees based on the distribution settings.
    
    Args:
        groups: List of experience groups
        assignee_distribution: Dict with assignee names as keys and counts as values
    
    Returns:
        List of groups with assignees assigned
    """
    # Filter out assignees with 0 count
    active_assignees = {name: count for name, count in assignee_distribution.items() if count > 0}
    
    if not active_assignees:
        return groups  # No assignees to assign
    
    # Create a list of assignees repeated by their count
    assignee_pool = []
    for name, count in active_assignees.items():
        assignee_pool.extend([name] * count)
    
    # Assign to groups (only those with non-empty experience IDs)
    assignee_idx = 0
    for group_key, group_size, group_rows in groups:
        # Check if this group has a valid experience ID
        exp_id = group_key[2]  # Associated Experience IDs is the third element
        has_exp_id = exp_id and str(exp_id).strip()
        
        if has_exp_id and assignee_pool:
            # Assign the same assignee to all rows in this group
            assignee = assignee_pool[assignee_idx % len(assignee_pool)]
            for row in group_rows:
                row['Assignee'] = assignee
            assignee_idx += 1
    
    return groups


def group_by_experience(rows):
    """Group rows by experience to identify which rows belong to the same experience."""
    groups = []
    current_group = []
    current_key = None
    
    for row in rows:
        # Create a key based on the three identifying columns
        key = (
            row.get('Associated Experience', ''),
            row.get('Backstage Experience page', ''),
            row.get('Associated Experience IDs', '')
        )
        
        # Check if key is all empty (don't group empty rows together)
        is_empty_key = all(k == '' for k in key)
        
        if current_key is None:
            # First row
            current_key = key
            current_group = [row]
        elif key == current_key and not is_empty_key:
            # Same experience AND not empty, add to current group
            current_group.append(row)
        else:
            # New experience OR empty key, save previous group and start new one
            groups.append((current_key, len(current_group), current_group))
            current_key = key
            current_group = [row]
    
    # Don't forget the last group
    if current_group:
        groups.append((current_key, len(current_group), current_group))
    
    return groups


def create_xlsx(rows, column_order, output_path):
    """Create XLSX file with proper formatting, merging cells for identical experiences."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Organized Tickets"
    
    # Define styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Alternating colors for experience groups
    colors = [
        PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid"),
        PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
    ]
    
    # Border styles
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    thick_border = Border(
        left=Side(style='medium', color='000000'),
        right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'),
        bottom=Side(style='medium', color='000000')
    )
    
    # Write header row
    for col_idx, col_name in enumerate(column_order, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thick_border
    
    # Group rows by experience
    groups = group_by_experience(rows)
    
    # Columns that should be merged for same experience
    merge_columns = ['Associated Experience', 'Backstage Experience page', 'Public Preview Link', 'Associated Experience IDs', 'Assignee']
    merge_col_indices = [column_order.index(col) + 1 for col in merge_columns]
    
    # Write data rows
    current_row = 2
    for group_idx, (experience_key, group_size, group_rows) in enumerate(groups):
        # Choose color for this experience group
        group_fill = colors[group_idx % len(colors)]
        
        start_row = current_row
        
        # Write all rows in this group
        for row_data in group_rows:
            for col_idx, col_name in enumerate(column_order, start=1):
                value = row_data[col_name]
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.fill = group_fill
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = thin_border
                
                # Add formula for Public Preview Link column
                if col_name == 'Public Preview Link':
                    # Find the column letter for Associated Experience IDs
                    exp_ids_col_idx = column_order.index('Associated Experience IDs') + 1
                    exp_ids_col_letter = get_column_letter(exp_ids_col_idx)
                    # Set formula: ="https://app.eko.com/public/experiences/" & D2
                    cell.value = f'="https://app.eko.com/public/experiences/" & {exp_ids_col_letter}{current_row}'
                
                # Make URLs clickable (no font styling)
                elif col_name == 'Backstage Experience page' and value and (str(value).startswith('http://') or str(value).startswith('https://')):
                    cell.hyperlink = str(value)
                
            current_row += 1
        
        # Merge cells for experience identification columns if group has multiple tickets
        if group_size > 1:
            end_row = current_row - 1
            for col_idx in merge_col_indices:
                ws.merge_cells(
                    start_row=start_row,
                    start_column=col_idx,
                    end_row=end_row,
                    end_column=col_idx
                )
                # Center the merged cell content
                merged_cell = ws.cell(row=start_row, column=col_idx)
                merged_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                merged_cell.border = thick_border
                
                # Make URLs clickable (no font styling)
                col_name = column_order[col_idx - 1]
                if col_name == 'Backstage Experience page' and merged_cell.value:
                    url = str(merged_cell.value)
                    if url.startswith('http://') or url.startswith('https://'):
                        merged_cell.hyperlink = url
    
    # Auto-adjust column widths
    for col_idx, col_name in enumerate(column_order, start=1):
        col_letter = get_column_letter(col_idx)
        
        # Set minimum and maximum widths
        if col_name in ['Ticket description', 'Ticket name']:
            ws.column_dimensions[col_letter].width = 40
        elif col_name in ['Associated Experience', 'Backstage Experience page', 'Public Preview Link']:
            ws.column_dimensions[col_letter].width = 30
        elif col_name in ['Associated Experience IDs', 'Ticket ID']:
            ws.column_dimensions[col_letter].width = 20
        elif col_name in ['Assignee', 'Ticket Category', 'Sub Category', 'Invalid ticket']:
            ws.column_dimensions[col_letter].width = 20
        elif col_name in ['Additional Notes', 'Notes - Admin', 'Non-ticketed issues']:
            ws.column_dimensions[col_letter].width = 30
        elif col_name in ['Next step', 'Last Updated - Status', 'Last Updated - Next step', 'Last Updated - Invalid Ticket', 'Last Updated - Non Ticketed Issues']:
            ws.column_dimensions[col_letter].width = 25
        else:
            ws.column_dimensions[col_letter].width = 15
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    max_row = len(rows) + 1
    
    # Add data validation dropdown for Ticket status column
    status_col_idx = column_order.index('Ticket status') + 1
    status_col_letter = get_column_letter(status_col_idx)
    
    dv_status = DataValidation(
        type="list",
        formula1='"Todo,In Progress,Done,Resolved,Stuck,Other ticket stuck"',
        allow_blank=True
    )
    dv_status.error = 'Please select a valid status from the dropdown'
    dv_status.errorTitle = 'Invalid Status'
    dv_status.prompt = 'Select a status from the dropdown'
    dv_status.promptTitle = 'Ticket Status'
    dv_status.add(f'{status_col_letter}2:{status_col_letter}{max_row}')
    ws.add_data_validation(dv_status)
    
    # Add data validation dropdown for Next step column (single select)
    next_step_col_idx = column_order.index('Next step') + 1
    next_step_col_letter = get_column_letter(next_step_col_idx)
    
    dv_next_step = DataValidation(
        type="list",
        formula1='"Reshoot,Manual post"',
        allow_blank=True
    )
    dv_next_step.prompt = 'Select one option'
    dv_next_step.promptTitle = 'Next Step'
    dv_next_step.add(f'{next_step_col_letter}2:{next_step_col_letter}{max_row}')
    ws.add_data_validation(dv_next_step)
    
    # Add data validation dropdown for Assignee column (single select)
    assignee_col_idx = column_order.index('Assignee') + 1
    assignee_col_letter = get_column_letter(assignee_col_idx)
    
    dv_assignee = DataValidation(
        type="list",
        formula1=f'"{",".join(ASSIGNEES)}"',
        allow_blank=True
    )
    dv_assignee.prompt = 'Select an assignee'
    dv_assignee.promptTitle = 'Assignee'
    dv_assignee.add(f'{assignee_col_letter}2:{assignee_col_letter}{max_row}')
    ws.add_data_validation(dv_assignee)
    
    # Save workbook
    wb.save(output_path)
    
    return groups


def process_csv_file(input_path, output_path, assignee_distribution=None):
    """Process CSV file and create organized XLSX output."""
    # Read CSV
    rows = read_csv_robust(input_path)
    
    # Process and reorder columns
    processed_rows, column_order = extract_and_reorder_columns(rows)
    
    # Group experiences
    groups = group_by_experience(processed_rows)
    
    # Assign experiences to assignees if distribution is provided
    if assignee_distribution:
        groups = assign_experiences_to_assignees(groups, assignee_distribution)
    
    # Create XLSX (we need to flatten groups back to rows for create_xlsx)
    all_rows = []
    for _, _, group_rows in groups:
        all_rows.extend(group_rows)
    
    # Create XLSX
    groups = create_xlsx(all_rows, column_order, output_path)
    
    # Return statistics
    multi_ticket_experiences = [g for g in groups if g[1] > 1]
    stats = {
        'total_tickets': len(all_rows),
        'unique_experiences': len(groups),
        'multi_ticket_experiences': len(multi_ticket_experiences)
    }
    
    return stats


@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and show assignee distribution page."""
    if 'file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a CSV file.', 'error')
        return redirect(url_for('index'))
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        # Read the CSV to count experiences
        rows = read_csv_robust(input_path)
        processed_rows, _ = extract_and_reorder_columns(rows)
        groups = group_by_experience(processed_rows)
        
        # Count experiences with valid IDs
        valid_experiences = sum(1 for key, _, _ in groups 
                               if key[2] and str(key[2]).strip())
        
        # Store filename in session
        session['uploaded_filename'] = filename
        session['total_experiences'] = valid_experiences
        
        # Redirect to assignee distribution page
        return redirect(url_for('assignee_distribution'))
        
    except Exception as e:
        flash(f'Error processing file: {str(e)}', 'error')
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)
        return redirect(url_for('index'))


@app.route('/assignee-distribution')
def assignee_distribution():
    """Show assignee distribution page."""
    if 'uploaded_filename' not in session:
        flash('No file uploaded', 'error')
        return redirect(url_for('index'))
    
    return render_template('assignee_distribution.html', 
                          assignees=ASSIGNEES,
                          total_experiences=session.get('total_experiences', 0))


@app.route('/process', methods=['POST'])
def process_file():
    """Process the file with assignee distribution."""
    if 'uploaded_filename' not in session:
        flash('No file uploaded', 'error')
        return redirect(url_for('index'))
    
    try:
        filename = session['uploaded_filename']
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        app.logger.info(f'Processing file: {filename}')
        app.logger.info(f'Input path: {input_path}')
        app.logger.info(f'Temp dir: {app.config["UPLOAD_FOLDER"]}')
        
        if not os.path.exists(input_path):
            app.logger.error(f'Input file not found at: {input_path}')
            flash('Uploaded file not found', 'error')
            return redirect(url_for('index'))
        
        # Get assignee distribution from request
        assignee_distribution = {}
        distribution_data = request.json if request.is_json else request.form
        
        for assignee in ASSIGNEES:
            count = int(distribution_data.get(f'assignee_{assignee}', 0))
            if count > 0:
                assignee_distribution[assignee] = count
        
        # Create output filename
        output_filename = Path(filename).stem + '_organized.xlsx'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        app.logger.info(f'Output path: {output_path}')
        
        # Process the file with assignee distribution
        stats = process_csv_file(input_path, output_path, assignee_distribution)
        
        app.logger.info(f'File processed successfully. Stats: {stats}')
        
        # Verify output file exists
        if not os.path.exists(output_path):
            app.logger.error(f'Output file was not created at: {output_path}')
            raise Exception('Failed to create output file')
        
        app.logger.info(f'Output file size: {os.path.getsize(output_path)} bytes')
        
        # Clean up input file
        os.remove(input_path)
        
        # Clear session
        session.pop('uploaded_filename', None)
        session.pop('total_experiences', None)
        
        # Set up cleanup after file is sent
        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                    app.logger.info(f'Cleaned up output file: {output_path}')
            except Exception as e:
                app.logger.error(f'Error removing file: {e}')
            return response
        
        # Send the file for download
        app.logger.info(f'Sending file for download: {output_filename}')
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        app.logger.error(f'Error in process_file: {str(e)}', exc_info=True)
        flash(f'Error processing file: {str(e)}', 'error')
        # Clean up files if they exist
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)
        if 'output_path' in locals() and os.path.exists(output_path):
            os.remove(output_path)
        session.pop('uploaded_filename', None)
        session.pop('total_experiences', None)
        return redirect(url_for('index'))


if __name__ == '__main__':
    # For local development
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=True, host='0.0.0.0', port=port)

