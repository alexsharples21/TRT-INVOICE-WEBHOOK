from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import mm
import base64
import io
import os
import re
import requests

app = Flask(__name__)

W, H = A4

# Microsoft Graph config — set these as environment variables on Railway
TENANT_ID     = os.environ.get('TENANT_ID', '')
CLIENT_ID     = os.environ.get('CLIENT_ID', '')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET', '')
ONEDRIVE_USER = os.environ.get('ONEDRIVE_USER', 'gary@trtmotorsltd.co.uk')
ONEDRIVE_FOLDER = os.environ.get('ONEDRIVE_FOLDER', 'TRT Invoices')

def get_access_token():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    data = {
        'grant_type':    'client_credentials',
        'client_id':     CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope':         'https://graph.microsoft.com/.default'
    }
    r = requests.post(url, data=data)
    return r.json().get('access_token', '')

def save_to_onedrive(pdf_bytes, filename):
    token = get_access_token()
    if not token:
        return {'success': False, 'error': 'Could not get access token'}
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/pdf'
    }
    # Upload to OneDrive — creates file if not exists, overwrites if it does
    url = f'https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER}/drive/root:/{ONEDRIVE_FOLDER}/{filename}:/content'
    r = requests.put(url, headers=headers, data=pdf_bytes)
    if r.status_code in [200, 201]:
        item = r.json()
        return {
            'success': True,
            'onedrive_url': item.get('webUrl', ''),
            'file_id': item.get('id', '')
        }
    return {'success': False, 'error': r.text, 'status': r.status_code}

def fmt(n):
    try:
        return f"\u00a3{float(n):,.2f}"
    except:
        return "\u00a30.00"

def extract_pdf_figures(pdf_base64):
    try:
        pdf_bytes = base64.b64decode(pdf_base64)
        text = pdf_bytes.decode('latin-1', errors='ignore')
    except:
        text = ''

    def extract_amount(pattern):
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return m.group(1).replace(',', '')
        return '0.00'

    labour     = extract_amount(r'Labour\s+\xa3([\d,]+\.\d{2})')
    parts      = extract_amount(r'Parts\s+\xa3([\d,]+\.\d{2})')
    paint      = extract_amount(r'Paints?\s*(?:&|and)\s*Materials\s+\xa3([\d,]+\.\d{2})')
    specialist = extract_amount(r'Specialist\s+\xa3([\d,]+\.\d{2})')
    grand      = extract_amount(r'Repairs Grand Total\s+\xa3([\d,]+\.\d{2})')
    not_vat    = bool(re.search(r'Not VAT registered', text, re.IGNORECASE))

    sub   = float(labour) + float(parts) + float(paint) + float(specialist)
    total = float(grand) if float(grand) > 0 else sub
    vat   = total - sub if not not_vat and total > sub else 0.0

    return {
        'labour':    f'{float(labour):.2f}',
        'parts':     f'{float(parts):.2f}',
        'paint':     f'{float(paint):.2f}',
        'specialist':f'{float(specialist):.2f}',
        'sub_total': f'{sub:.2f}',
        'vat':       f'{vat:.2f}',
        'grand_total':f'{total:.2f}',
        'vat_registered': not not_vat
    }

def generate_invoice(data):
    name        = data.get('name', '')
    address     = data.get('address', '')
    reg         = data.get('reg', '')
    vehicle     = data.get('vehicle', '').upper()
    inv_num     = data.get('invoice_number', 'TBC')
    acg_ref     = data.get('acg_ref', '')
    date        = data.get('date', '')
    labour      = data.get('labour', '0.00')
    parts       = data.get('parts', '0.00')
    paint       = data.get('paint', '0.00')
    specialist  = data.get('specialist', '0.00')
    sub_total   = data.get('sub_total', '0.00')
    vat         = data.get('vat', '0.00')
    grand_total = data.get('grand_total', '0.00')

    addr_lines = [l.strip() for l in address.split(',') if l.strip()]

    table_data = [
        ['LABOUR',              fmt(labour)],
        ['PARTS',               fmt(parts)],
        ['PAINT AND MATERIALS', fmt(paint)],
        ['SPECIALIST',          fmt(specialist)],
        ['SUB TOTAL',           fmt(sub_total)],
        ['VAT',                 fmt(vat)],
        ['GRAND TOTAL',         fmt(grand_total)],
    ]

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    c.setFont('Helvetica', 7.5)
    c.setFillColor(colors.HexColor('#333333'))
    c.drawCentredString(W/2, H-18*mm, '39/40 Fox Street Liverpool, L3 3BQ  Tel: 0151 207 9817  |  Email: trtmotors@hotmail.co.uk')
    c.drawCentredString(W/2, H-22*mm, 'Company No: 8372128  Vat Reg No: 320371937')

    logo_path = '/app/trt_logo_clean.png'
    logo_w, logo_h = 55*mm, 46*mm
    if os.path.exists(logo_path):
        c.drawImage(logo_path, (W-logo_w)/2, H-22*mm-logo_h-5*mm,
                    width=logo_w, height=logo_h, preserveAspectRatio=True)

    y = H - 22*mm - logo_h - 12*mm
    left, right = 25*mm, W - 25*mm

    c.setFont('Helvetica', 10)
    c.setFillColor(colors.black)
    c.drawString(left, y, name)
    c.drawRightString(right, y, date)

    for line in addr_lines:
        y -= 5*mm
        c.drawString(left, y, line)

    y -= 6*mm
    c.drawString(left, y, f'Invoice Number {inv_num}')
    y -= 5*mm
    c.drawString(left, y, f'REF {acg_ref}')
    y -= 8*mm
    c.drawCentredString(W/2, y, f'{vehicle}  REG {reg}')
    y -= 10*mm

    col_w = [110*mm, 50*mm]
    tbl = Table(table_data, colWidths=col_w, rowHeights=7*mm)
    tbl.setStyle(TableStyle([
        ('FONT',     (0,0),(-1,-1), 'Helvetica',     10),
        ('FONT',     (0,0),(0,-1),  'Helvetica-Bold', 10),
        ('FONT',     (0,4),(-1,-1), 'Helvetica-Bold', 10),
        ('ALIGN',    (0,0),(0,-1),  'RIGHT'),
        ('ALIGN',    (1,0),(1,-1),  'LEFT'),
        ('VALIGN',   (0,0),(-1,-1), 'MIDDLE'),
        ('BOX',      (0,0),(-1,-1), 0.5, colors.black),
        ('INNERGRID',(0,0),(-1,-1), 0.5, colors.black),
        ('LEFTPADDING', (0,0),(-1,-1), 6),
        ('RIGHTPADDING',(0,0),(-1,-1), 6),
    ]))
    tbl_w = sum(col_w)
    tbl.wrapOn(c, tbl_w, 200*mm)
    tbl_h = tbl._height
    tbl.drawOn(c, (W-tbl_w)/2, y - tbl_h)
    y = y - tbl_h - 12*mm

    c.setFont('Helvetica-Bold', 10)
    c.setFillColor(colors.red)
    c.drawString(left, y, 'PLEASE MAKE PAYMENTS TO:')
    y -= 5*mm
    c.setFont('Helvetica', 10)
    c.setFillColor(colors.black)
    for line in ['TRT MOTORS LTD','NATWEST','ACCOUNT NUMBER: 75014610','SORT CODE: 537021']:
        c.drawString(left, y, line)
        y -= 5*mm

    for fy, fh, fs, ft in [
        (12*mm, 8*mm, 7.5, '39/40 Fox Street Liverpool, L3 3BQ  Tel: 0151 207 9817  |  Email: trtmotors@hotmail.co.uk'),
        (6*mm,  5*mm, 7.0, 'Company No: 8372128  Vat Reg No: 320371937')
    ]:
        c.setFillColor(colors.HexColor('#1a1aff'))
        c.rect(20*mm, fy, W-40*mm, fh, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont('Helvetica', fs)
        c.drawCentredString(W/2, fy + fh*0.35, ft)

    c.save()
    buffer.seek(0)
    return buffer.read()


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/extract-and-generate', methods=['POST'])
def extract_and_generate():
    data = request.json or {}
    pdf_base64 = data.get('pdf_base64', '')

    figures = extract_pdf_figures(pdf_base64) if pdf_base64 else {
        'labour':'0.00','parts':'0.00','paint':'0.00',
        'specialist':'0.00','sub_total':'0.00','vat':'0.00','grand_total':'0.00'
    }

    invoice_data = {**data, **figures}

    try:
        pdf_bytes = generate_invoice(invoice_data)
        reg_clean = data.get('reg','').replace(' ','')
        inv_num   = data.get('invoice_number','TBC')
        filename  = f"TRT_Invoice_{inv_num}_{reg_clean}.pdf"

        # Save to OneDrive
        onedrive_result = save_to_onedrive(pdf_bytes, filename)

        return jsonify({
            'figures':      figures,
            'pdf_base64':   base64.b64encode(pdf_bytes).decode('utf-8'),
            'filename':     filename,
            'onedrive':     onedrive_result,
            'success':      True
        })
    except Exception as e:
        return jsonify({'error': str(e), 'figures': figures, 'success': False}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
