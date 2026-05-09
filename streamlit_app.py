import io
import os
from datetime import datetime
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Pt
from fpdf import FPDF


def format_date_for_export(date_str):
    if not date_str:
        return ''
    try:
        return datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return date_str


def create_word_bytes(date, time, attendees, points, action_items, meeting_heading=''):
    doc = Document()

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Verdana'
    normal_style.font.size = Pt(10)

    heading_style = doc.styles['Heading 1']
    heading_style.font.name = 'Verdana'
    heading_style.font.size = Pt(11)
    heading_style.font.bold = True

    title_para = doc.add_paragraph()
    title_run = title_para.add_run('Minutes of Meeting')
    title_run.font.name = 'Verdana'
    title_run.font.size = Pt(11)
    title_run.font.bold = True

    mh = meeting_heading.strip() if meeting_heading else ''
    if mh:
        heading_para = doc.add_paragraph()
        heading_run = heading_para.add_run(mh)
        heading_run.font.name = 'Verdana'
        heading_run.font.size = Pt(11)
        heading_run.font.bold = True

    doc.add_paragraph(f'Date: {format_date_for_export(date)}')
    doc.add_paragraph(f'Time: {time}')

    doc.add_heading('Members Present', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'S.No'
    hdr_cells[1].text = 'Name'
    hdr_cells[2].text = 'Designation / Role'

    sno = 1
    for attendee in attendees:
        if attendee['name'].strip():
            row_cells = table.add_row().cells
            row_cells[0].text = str(sno)
            row_cells[1].text = attendee['name']
            row_cells[2].text = attendee['role']
            sno += 1

    doc.add_heading('Points Discussed', level=1)
    doc.add_paragraph(points if points else '')

    if action_items:
        doc.add_heading('Action Items', level=1)
        action_table = doc.add_table(rows=1, cols=3)
        action_table.style = 'Light Grid Accent 1'
        action_hdr = action_table.rows[0].cells
        action_hdr[0].text = 'Action Item'
        action_hdr[1].text = 'Responsibility'
        action_hdr[2].text = 'Timeline'

        for item in action_items:
            action_row = action_table.add_row().cells
            action_row[0].text = item['desc']
            action_row[1].text = item['resp']
            action_row[2].text = item['date']

    doc.add_heading('Signatures', level=1)
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
    for attendee in attendees:
        if attendee['name'].strip():
            if sno % 2 == 1:
                if sno > 1:
                    sig_row = sig_table.add_row()
                else:
                    sig_row = sig_table.rows[0]
            else:
                sig_row = sig_table.rows[-1]

            cell_idx = (sno - 1) % 2
            sig_para = sig_row.cells[cell_idx].paragraphs[0]
            sig_para.alignment = 1
            sig_run = sig_para.add_run('_' * 20 + '\n' + attendee['name'] + '\n' + attendee['role'])
            sig_run.font.name = 'Verdana'
            sig_run.font.size = Pt(10)
            sno += 1

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def create_pdf_bytes(date, time, attendees, points, action_items, meeting_heading=''):
    pdf = FPDF()
    font_family = 'Arial'

    windows_fonts = {
        '': 'C:/Windows/Fonts/verdana.ttf',
        'B': 'C:/Windows/Fonts/verdanab.ttf',
        'I': 'C:/Windows/Fonts/verdanai.ttf',
    }

    try:
        for style_key, font_path in windows_fonts.items():
            if os.path.exists(font_path):
                pdf.add_font('Verdana', style_key, font_path, uni=True)
        if all(os.path.exists(path) for path in windows_fonts.values()):
            font_family = 'Verdana'
    except Exception:
        font_family = 'Arial'

    pdf.add_page()
    pdf.set_font(font_family, 'B', 11)
    pdf.cell(0, 10, txt='Minutes of Meeting', ln=True, align='C')
    pdf.set_font(font_family, '', 10)
    pdf.ln(3)
    mh = meeting_heading.strip() if meeting_heading else ''
    if mh:
        pdf.set_font(font_family, 'B', 11)
        pdf.multi_cell(0, 6, mh, align='C')
        pdf.ln(2)
    pdf.set_font(font_family, '', 10)
    pdf.cell(0, 8, txt=f'Date: {format_date_for_export(date)}', ln=True)
    pdf.cell(0, 8, txt=f'Time: {time}', ln=True)
    pdf.ln(5)

    pdf.set_font(font_family, 'B', 11)
    pdf.cell(0, 8, txt='Members Present:', ln=True)
    pdf.set_font(font_family, '', 10)

    col_widths = [15, 80, 95]
    pdf.cell(col_widths[0], 7, 'S.No', border=1, align='C')
    pdf.cell(col_widths[1], 7, 'Name', border=1, align='C')
    pdf.cell(col_widths[2], 7, 'Designation / Role', border=1, align='C', ln=True)

    sno = 1
    for attendee in attendees:
        if attendee['name'].strip():
            pdf.cell(col_widths[0], 7, str(sno), border=1, align='C')
            pdf.cell(col_widths[1], 7, attendee['name'][:30], border=1)
            pdf.cell(col_widths[2], 7, attendee['role'][:35], border=1, ln=True)
            sno += 1
    pdf.ln(5)

    pdf.set_font(font_family, 'B', 11)
    pdf.cell(0, 8, txt='Points Discussed:', ln=True)
    pdf.set_font(font_family, '', 10)
    pdf.multi_cell(0, 5, points if points else '')
    pdf.ln(5)

    if action_items:
        pdf.set_font(font_family, 'B', 11)
        pdf.cell(0, 8, txt='Action Items:', ln=True)
        pdf.set_font(font_family, '', 10)

        action_col_widths = [65, 65, 40]
        pdf.cell(action_col_widths[0], 7, 'Action Item', border=1, align='C')
        pdf.cell(action_col_widths[1], 7, 'Responsibility', border=1, align='C')
        pdf.cell(action_col_widths[2], 7, 'Timeline', border=1, align='C', ln=True)

        for item in action_items:
            desc = item['desc']
            resp = item['resp']
            date_timeline = item['date']
            max_desc_lines = max(1, len(desc) // 45 + 1)
            max_resp_lines = max(1, len(resp) // 35 + 1)
            max_lines = max(max_desc_lines, max_resp_lines, 1)
            row_height = max(max_lines * 4, 8)

            x_pos = pdf.get_x()
            y_pos = pdf.get_y()

            pdf.set_xy(x_pos, y_pos)
            pdf.set_draw_color(200, 200, 200)
            pdf.rect(x_pos, y_pos, action_col_widths[0], row_height, 'D')
            pdf.set_xy(x_pos + 1, y_pos + 1)
            pdf.set_draw_color(0, 0, 0)
            pdf.multi_cell(action_col_widths[0] - 2, 3.5, desc)

            pdf.set_xy(x_pos + action_col_widths[0], y_pos)
            pdf.rect(x_pos + action_col_widths[0], y_pos, action_col_widths[1], row_height, 'D')
            pdf.set_xy(x_pos + action_col_widths[0] + 1, y_pos + 1)
            pdf.multi_cell(action_col_widths[1] - 2, 3.5, resp)

            pdf.set_xy(x_pos + action_col_widths[0] + action_col_widths[1], y_pos)
            pdf.rect(x_pos + action_col_widths[0] + action_col_widths[1], y_pos, action_col_widths[2], row_height, 'D')
            pdf.set_xy(x_pos + action_col_widths[0] + action_col_widths[1] + 1, y_pos + 1)
            pdf.multi_cell(action_col_widths[2] - 2, 3.5, date_timeline)

            pdf.set_xy(x_pos, y_pos + row_height)
        pdf.ln(8)

    pdf.set_font(font_family, 'B', 11)
    pdf.cell(0, 8, txt='Signatures:', ln=True)
    pdf.set_font(font_family, 'I', 10)
    pdf.cell(0, 6, txt='Agreed and acknowledged by the members present:', ln=True)
    pdf.ln(3)

    sig_per_row = 2
    sig_x_start = 15
    sig_width = 90
    sig_y = pdf.get_y()

    for idx, attendee in enumerate(attendees):
        if attendee['name'].strip():
            col_idx = idx % sig_per_row
            if col_idx == 0 and idx > 0:
                sig_y += 25

            x_pos = sig_x_start + (col_idx * sig_width)
            pdf.set_xy(x_pos, sig_y)
            pdf.cell(sig_width, 2, '_' * 30, border=0, align='C')
            pdf.set_xy(x_pos, sig_y + 3)
            pdf.cell(sig_width, 5, attendee['name'], border=0, align='C')
            pdf.set_xy(x_pos, sig_y + 7)
            pdf.cell(sig_width, 4, attendee['role'], border=0, align='C')

    pdf_bytes = pdf.output(dest='S').encode('latin-1')
    return pdf_bytes


def ensure_state():
    if 'meeting_heading' not in st.session_state:
        st.session_state.meeting_heading = ''
    if 'attendees' not in st.session_state:
        st.session_state.attendees = [{'id': 1, 'name': '', 'role': ''}]
    if 'actions' not in st.session_state:
        st.session_state.actions = [{'id': 1, 'desc': '', 'resp': '', 'date': ''}]
    if 'attendee_counter' not in st.session_state:
        st.session_state.attendee_counter = 2
    if 'action_counter' not in st.session_state:
        st.session_state.action_counter = 2


def add_attendee():
    st.session_state.attendees.append({'id': st.session_state.attendee_counter, 'name': '', 'role': ''})
    st.session_state.attendee_counter += 1


def add_action():
    st.session_state.actions.append({'id': st.session_state.action_counter, 'desc': '', 'resp': '', 'date': ''})
    st.session_state.action_counter += 1


def remove_attendee(attendee_id):
    st.session_state.attendees = [a for a in st.session_state.attendees if a['id'] != attendee_id]


def remove_action(action_id):
    st.session_state.actions = [a for a in st.session_state.actions if a['id'] != action_id]


def load_css():
    css_path = Path(__file__).parent / 'styles' / 'app.css'
    if css_path.exists():
        st.markdown(f'<style>{css_path.read_text(encoding="utf-8")}</style>', unsafe_allow_html=True)


def main():
    st.set_page_config(page_title='MOM Creator', layout='wide', initial_sidebar_state='collapsed')
    load_css()
    
    st.title('Minutes of Meeting (MOM) Creator')
    st.caption('Create professional MOM documents in Word and PDF format.')

    ensure_state()

    with st.expander('📅 Meeting Details', expanded=True):
        st.text_input(
            'Meeting Heading',
            key='meeting_heading',
            placeholder='Subject or title shown on exported documents',
            help='Printed below “Minutes of Meeting” in Word and PDF.',
        )
        cols = st.columns(2)
        with cols[0]:
            with st.container(border=True):
                date = st.date_input('Meeting Date', help='Select the date of the meeting')
        with cols[1]:
            with st.container(border=True):
                time_value = st.time_input('Meeting Time', help='Select meeting start time')
                time = time_value.strftime('%I:%M %p')

    st.markdown(
        '<h2 class="section-heading-18">Attendees</h2>',
        unsafe_allow_html=True,
    )
    
    b1, b2, _ = st.columns([2, 2, 6], gap='small')
    with b1:
        if st.button(
            'BTN_ADD_ATTENDEE',
            key='add_attendee_btn',
            use_container_width=True,
            help='Add a row (Ctrl+Shift+1)',
        ):
            add_attendee()
    with b2:
        if st.button(
            'BTN_RESET_ATTENDEE',
            key='reset_attendee_btn',
            use_container_width=True,
            help='Clear attendee list (Ctrl+Shift+2)',
        ):
            st.session_state.attendees = [{'id': 1, 'name': '', 'role': ''}]
            st.session_state.attendee_counter = 2
            st.rerun()

    with st.container(border=True):
        header_cols = st.columns([4, 4, 1])
        header_cols[0].markdown('<div class="row-header">Name</div>', unsafe_allow_html=True)
        header_cols[1].markdown('<div class="row-header">Role / Designation</div>', unsafe_allow_html=True)
        header_cols[2].markdown('<div class="row-header">Remove</div>', unsafe_allow_html=True)

        for idx, attendee in enumerate(st.session_state.attendees, 1):
            cols = st.columns([4, 4, 1])
            attendee['name'] = cols[0].text_input(
                f'Attendee Name {idx}',
                value=attendee['name'],
                key=f"attendee_name_{attendee['id']}",
                placeholder='Enter full name',
                label_visibility='collapsed'
            )
            attendee['role'] = cols[1].text_input(
                f'Attendee Role {idx}',
                value=attendee['role'],
                key=f"attendee_role_{attendee['id']}",
                placeholder='Enter role/designation',
                label_visibility='collapsed'
            )
            if cols[2].button('X', key=f"remove_attendee_{attendee['id']}"):
                remove_attendee(attendee['id'])
                st.rerun()

    st.markdown(
        '<h2 class="section-heading-18">Points Discussed</h2>',
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        points = st.text_area('Enter points discussed', height=200, placeholder='Type or paste the points discussed during the meeting...')

    st.markdown(
        '<h2 class="section-heading-18">Action Items</h2>',
        unsafe_allow_html=True,
    )
    
    a1, a2, _ = st.columns([2, 2, 6], gap='small')
    with a1:
        if st.button(
            'BTN_ADD_ACTION',
            key='add_action_btn',
            use_container_width=True,
            help='Add an action row (Ctrl+Shift+3)',
        ):
            add_action()
    with a2:
        if st.button(
            'BTN_RESET_ACTION',
            key='reset_action_btn',
            use_container_width=True,
            help='Clear action items (Ctrl+Shift+4)',
        ):
            st.session_state.actions = [{'id': 1, 'desc': '', 'resp': '', 'date': ''}]
            st.session_state.action_counter = 2
            st.rerun()
    
    with st.container(border=True):
        action_header_cols = st.columns([5, 4, 3, 1])
        action_header_cols[0].markdown('<div class="row-header">Action Item</div>', unsafe_allow_html=True)
        action_header_cols[1].markdown('<div class="row-header">Responsibility</div>', unsafe_allow_html=True)
        action_header_cols[2].markdown('<div class="row-header">Timeline</div>', unsafe_allow_html=True)
        action_header_cols[3].markdown('<div class="row-header">Remove</div>', unsafe_allow_html=True)

        for idx, action in enumerate(st.session_state.actions, 1):
            cols = st.columns([5, 4, 3, 1])
            action['desc'] = cols[0].text_input(
                f'Action Item Description {idx}',
                value=action['desc'],
                key=f"action_desc_{action['id']}",
                placeholder='What needs to be done?',
                label_visibility='collapsed'
            )
            action['resp'] = cols[1].text_input(
                f'Action Item Owner {idx}',
                value=action['resp'],
                key=f"action_resp_{action['id']}",
                placeholder='Who is responsible?',
                label_visibility='collapsed'
            )
            action['date'] = cols[2].text_input(
                f'Action Item Timeline {idx}',
                value=action['date'],
                key=f"action_date_{action['id']}",
                placeholder='When is it due?',
                label_visibility='collapsed'
            )
            if cols[3].button('X', key=f"remove_action_{action['id']}"):
                remove_action(action['id'])
                st.rerun()

    attendees = [a for a in st.session_state.attendees if a['name'].strip()]
    action_items = [a for a in st.session_state.actions if a['desc'].strip()]

    st.markdown(
        '<h2 class="section-heading-18">Download Documents</h2>',
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        # Two narrow cols + spacer so buttons sit side-by-side; wide col eats flex space only on the right
        dl1, dl2, _ = st.columns([2, 2, 6], gap='small')
        with dl1:
            word_bytes = create_word_bytes(
                date.isoformat(),
                time,
                attendees,
                points,
                action_items,
                meeting_heading=st.session_state.get('meeting_heading', ''),
            )
            st.download_button(
                'BTN_DOWNLOAD_WORD',
                word_bytes,
                file_name='minutes_of_meeting.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                help='Download Word (Ctrl+Shift+W)',
                key='download_word',
                use_container_width=True,
            )
        with dl2:
            pdf_bytes = create_pdf_bytes(
                date.isoformat(),
                time,
                attendees,
                points,
                action_items,
                meeting_heading=st.session_state.get('meeting_heading', ''),
            )
            st.download_button(
                'BTN_DOWNLOAD_PDF',
                pdf_bytes,
                file_name='minutes_of_meeting.pdf',
                mime='application/pdf',
                help='Download PDF (Ctrl+Shift+P)',
                key='download_pdf',
                use_container_width=True,
            )
    with st.expander('Keyboard shortcuts', expanded=False):
        st.markdown(
            """
| Control | Shortcut |
|---------|----------|
| Add attendee row | Ctrl+Shift+1 |
| Reset attendees | Ctrl+Shift+2 |
| Add action row | Ctrl+Shift+3 |
| Reset action items | Ctrl+Shift+4 |
| Download Word | Ctrl+Shift+W |
| Download PDF | Ctrl+Shift+P |
"""
        )
    st.info('Tip: For Streamlit Community Cloud deployment, set `streamlit_app.py` as the main file.')


if __name__ == '__main__':
    main()
