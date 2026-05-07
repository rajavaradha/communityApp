from flask import Flask, render_template, request, send_file
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt
from fpdf import FPDF

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def format_date_for_export(date_str):
    if not date_str:
        return ''
    try:
        return datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return date_str

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date = request.form.get('date')
        time = request.form.get('time')
        attendees_names = request.form.getlist('attendee_name[]')
        attendees_roles = request.form.getlist('attendee_role[]')
        points = request.form.get('points')
        action_descs = request.form.getlist('action_desc[]')
        action_dates = request.form.getlist('action_date[]')
        action_resps = request.form.getlist('action_resp[]')
        formatted_date = format_date_for_export(date)

        if 'export_word' in request.form:
            action_items = [(desc, resp, adate) for desc, resp, adate in zip(action_descs, action_resps, action_dates) if desc.strip()]
            
            doc = Document()
            
            # Set default font for the document
            style = doc.styles['Normal']
            style.font.name = 'Verdana'
            style.font.size = Pt(10)
            
            title_style = doc.styles.add_style('CustomTitle', 1)
            title_style.font.name = 'Verdana'
            title_style.font.size = Pt(11)
            title_style.font.bold = True
            
            heading_style = doc.styles.add_style('CustomHeading', 1)
            heading_style.font.name = 'Verdana'
            heading_style.font.size = Pt(11)
            heading_style.font.bold = True
            
            doc.add_heading('Minutes of Meeting', 0).style = title_style
            doc.add_paragraph(f"Date: {formatted_date}")
            doc.add_paragraph(f"Time: {time}")
            
            # Members Present Table
            doc.add_heading('Members Present', level=1).style = heading_style
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Light Grid Accent 1'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'S.No'
            hdr_cells[1].text = 'Name'
            hdr_cells[2].text = 'Designation / Role'
            
            sno = 1
            for name, role in zip(attendees_names, attendees_roles):
                if name.strip():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(sno)
                    row_cells[1].text = name
                    row_cells[2].text = role
                    sno += 1
            
            # Points Discussed
            doc.add_heading('Points Discussed', level=1).style = heading_style
            doc.add_paragraph(points if points else "")
            
            if action_items:
                # Action Items Table
                doc.add_heading('Action Items', level=1).style = heading_style
                action_table = doc.add_table(rows=1, cols=3)
                action_table.style = 'Light Grid Accent 1'
                action_hdr = action_table.rows[0].cells
                action_hdr[0].text = 'Action Item'
                action_hdr[1].text = 'Responsibility'
                action_hdr[2].text = 'Timeline'
                
                for desc, resp, adate in action_items:
                    action_row = action_table.add_row().cells
                    action_row[0].text = desc
                    action_row[1].text = resp
                    action_row[2].text = adate
            
            # Signatures
            doc.add_heading('Signatures', level=1).style = heading_style
            para = doc.add_paragraph()
            run = para.add_run('Agreed and acknowledged by the members present:')
            run.font.name = 'Verdana'
            run.font.size = Pt(10)
            run.italic = True
            doc.add_paragraph()
            
            sig_table = doc.add_table(rows=1, cols=2)
            sig_table.autofit = False
            sig_table.allow_autofit = False
            
            sno = 1
            for name, role in zip(attendees_names, attendees_roles):
                if name.strip():
                    if sno % 2 == 1:
                        if sno > 1:
                            sig_row = sig_table.add_row()
                        else:
                            sig_row = sig_table.rows[0]
                    else:
                        sig_row = sig_table.rows[-1]
                    
                    cell_idx = (sno - 1) % 2
                    sig_para = sig_row.cells[cell_idx].paragraphs[0]
                    sig_para.alignment = 1  # Center
                    sig_run = sig_para.add_run('_' * 20 + '\n' + name + '\n' + role)
                    sig_run.font.name = 'Verdana'
                    sig_run.font.size = Pt(10)
                    sno += 1
            
            filename = os.path.join(app.config['UPLOAD_FOLDER'], 'mom.docx')
            doc.save(filename)
            return send_file(filename, as_attachment=True, download_name='mom.docx')

        elif 'export_pdf' in request.form:
            action_items = [(desc, resp, date_timeline) for desc, resp, date_timeline in zip(action_descs, action_resps, action_dates) if desc.strip()]

            pdf = FPDF()
            # Add Verdana font if available
            try:
                pdf.add_font('Verdana', '', 'C:/Windows/Fonts/verdana.ttf', uni=True)
                pdf.add_font('Verdana', 'B', 'C:/Windows/Fonts/verdanab.ttf', uni=True)
                pdf.add_font('Verdana', 'I', 'C:/Windows/Fonts/verdanai.ttf', uni=True)
                font_family = 'Verdana'
            except:
                font_family = 'Arial'  # Fallback to Arial if Verdana not available
            
            pdf.add_page()
            pdf.set_font(font_family, 'B', size=11)
            pdf.cell(0, 10, txt="Minutes of Meeting", ln=True, align='C')
            pdf.set_font(font_family, size=10)
            pdf.ln(5)
            pdf.cell(0, 8, txt=f"Date: {formatted_date}", ln=True)
            pdf.cell(0, 8, txt=f"Time: {time}", ln=True)
            pdf.ln(5)
            
            # Members Present Table
            pdf.set_font(font_family, 'B', size=11)
            pdf.cell(0, 8, txt="Members Present:", ln=True)
            pdf.set_font(font_family, size=10)
            
            col_widths = [15, 80, 95]
            pdf.cell(col_widths[0], 7, "S.No", border=1, align='C')
            pdf.cell(col_widths[1], 7, "Name", border=1, align='C')
            pdf.cell(col_widths[2], 7, "Designation / Role", border=1, align='C', ln=True)
            
            sno = 1
            for name, role in zip(attendees_names, attendees_roles):
                if name.strip():
                    pdf.cell(col_widths[0], 7, str(sno), border=1, align='C')
                    pdf.cell(col_widths[1], 7, name[:30], border=1)
                    pdf.cell(col_widths[2], 7, role[:35], border=1, ln=True)
                    sno += 1
            pdf.ln(5)
            
            # Points Discussed
            pdf.set_font(font_family, 'B', size=11)
            pdf.cell(0, 8, txt="Points Discussed:", ln=True)
            pdf.set_font(font_family, size=10)
            pdf.multi_cell(0, 5, points if points else "")
            pdf.ln(5)
            
            if action_items:
                # Action Items Table
                pdf.set_font(font_family, 'B', size=11)
                pdf.cell(0, 8, txt="Action Items:", ln=True)
                pdf.set_font(font_family, size=10)
                
                action_col_widths = [65, 65, 40]
                pdf.cell(action_col_widths[0], 7, "Action Item", border=1, align='C')
                pdf.cell(action_col_widths[1], 7, "Responsibility", border=1, align='C')
                pdf.cell(action_col_widths[2], 7, "Timeline", border=1, align='C', ln=True)
                
                for desc, resp, date_timeline in action_items:
                    # Calculate row height based on text length
                    max_desc_lines = max(1, len(desc) // 45 + 1)
                    max_resp_lines = max(1, len(resp) // 35 + 1)
                    max_lines = max(max_desc_lines, max_resp_lines, 1)
                    row_height = max(max_lines * 4, 8)
                    
                    # Store current position to create multi-column cells
                    x_pos = pdf.get_x()
                    y_pos = pdf.get_y()
                    
                    # Action Item column (wrapping)
                    pdf.set_xy(x_pos, y_pos)
                    pdf.set_draw_color(200, 200, 200)
                    pdf.rect(x_pos, y_pos, action_col_widths[0], row_height, 'D')
                    pdf.set_xy(x_pos + 1, y_pos + 1)
                    pdf.set_draw_color(0, 0, 0)
                    pdf.multi_cell(action_col_widths[0] - 2, 3.5, desc)
                    
                    # Responsibility column
                    pdf.set_xy(x_pos + action_col_widths[0], y_pos)
                    pdf.rect(x_pos + action_col_widths[0], y_pos, action_col_widths[1], row_height, 'D')
                    pdf.set_xy(x_pos + action_col_widths[0] + 1, y_pos + 1)
                    pdf.multi_cell(action_col_widths[1] - 2, 3.5, resp)
                    
                    # Timeline column
                    pdf.set_xy(x_pos + action_col_widths[0] + action_col_widths[1], y_pos)
                    pdf.rect(x_pos + action_col_widths[0] + action_col_widths[1], y_pos, action_col_widths[2], row_height, 'D')
                    pdf.set_xy(x_pos + action_col_widths[0] + action_col_widths[1] + 1, y_pos + 1)
                    pdf.multi_cell(action_col_widths[2] - 2, 3.5, date_timeline)
                    
                    # Move to next row
                    pdf.set_xy(x_pos, y_pos + row_height)
                pdf.ln(8)
            
            # Signatures
            pdf.set_font(font_family, 'B', size=11)
            pdf.cell(0, 8, txt="Signatures:", ln=True)
            pdf.set_font(font_family, 'I', size=10)
            pdf.cell(0, 6, txt="Agreed and acknowledged by the members present:", ln=True)
            pdf.ln(3)
            
            pdf.set_font(font_family, size=10)
            sig_count = 0
            for name, role in zip(attendees_names, attendees_roles):
                if name.strip():
                    sig_count += 1
            
            sig_per_row = 2
            sig_x_start = 15
            sig_width = 90
            sig_y = pdf.get_y()
            
            for idx, (name, role) in enumerate(zip(attendees_names, attendees_roles)):
                if name.strip():
                    col_idx = idx % sig_per_row
                    if col_idx == 0 and idx > 0:
                        sig_y += 25
                    
                    x_pos = sig_x_start + (col_idx * sig_width)
                    pdf.set_xy(x_pos, sig_y)
                    pdf.cell(sig_width, 2, "_" * 30, border=0, align='C')
                    pdf.set_xy(x_pos, sig_y + 3)
                    pdf.cell(sig_width, 5, name, border=0, align='C')
                    pdf.set_xy(x_pos, sig_y + 7)
                    pdf.cell(sig_width, 4, role, border=0, align='C')
            
            filename = os.path.join(app.config['UPLOAD_FOLDER'], 'mom.pdf')
            pdf.output(filename)
            return send_file(filename, as_attachment=True, download_name='mom.pdf')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)