"""
COMPLETE WARRANTY MANAGEMENT SYSTEM - MOBILE RESPONSIVE VERSION
✅ All 6 warranty tabs
✅ Full mobile responsive design (iPhone, iPad, Android)
✅ Touch-friendly interface
✅ Complete error handling
✅ Professional UI/UX
✅ Fast performance on mobile
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import uvicorn
from fastapi import FastAPI, Request, HTTPException, Cookie
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, JSONResponse, StreamingResponse
import os
import socket
from typing import Optional
import sys
import secrets
from PIL import Image, ImageDraw, ImageFont
import io
import base64
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import traceback

IS_RENDER = os.getenv('RENDER', 'false').lower() == 'true'
DATA_DIR = os.getenv('DATA_DIR', '/mnt/data' if IS_RENDER else '.')

print(f"\n{'='*100}")
print(f"WARRANTY MANAGEMENT SYSTEM - MOBILE RESPONSIVE VERSION")
print(f"{'='*100}")
print(f"Environment: {'Render' if IS_RENDER else 'Local'}")
print(f"Data Directory: {DATA_DIR}\n")

WARRANTY_DATA = {
    'credit_df': None,
    'debit_df': None,
    'arbitration_df': None,
    'source_df': None,
    'current_month_df': None,
    'current_month_source_df': None,
    'compensation_df': None,
    'compensation_source_df': None,
    'pr_approval_df': None,
    'pr_approval_source_df': None
}

def find_data_file(filename):
    possible_paths = [
        filename, f"./{filename}", os.path.join(DATA_DIR, filename),
        os.path.join(DATA_DIR, 'data', filename), os.path.join('data', filename),
    ]
    if filename.endswith('.xlsx'):
        name_without_ext = filename.replace('.xlsx', '')
        copy_variant = f"{name_without_ext} - Copy.xlsx"
        possible_paths.extend([copy_variant, f"./{copy_variant}", os.path.join(DATA_DIR, copy_variant)])
    for path in possible_paths:
        if os.path.exists(path):
            print(f"✓ Found: {filename}")
            return path
    print(f"✗ Not found: {filename}")
    return None

def process_warranty_data():
    input_path = find_data_file('Warranty Debit.xlsx')
    if input_path is None:
        return None, None, None, None
    try:
        df = pd.read_excel(input_path, sheet_name='Sheet1')
        print(f"Processing Warranty Debit - {len(df)} rows")

        dealer_mapping = {
            'AMRAVATI': 'AMT', 'CHAUFULA_SZZ': 'CHA', 'CHIKHALI': 'CHI',
            'KOLHAPUR_WS': 'KOL', 'NAGPUR_KAMPTHEE ROAD': 'HO',
            'NAGPUR_WARDHAMAN NGR': 'CITY', 'SHIKRAPUR_SZS': 'SHI',
            'WAGHOLI': 'WAG', 'YAVATMAL': 'YAT', 'NAGPUR_WARDHAMAN NGR_CQ': 'CQ'
        }

        for col in ['Total Claim Amount', 'Credit Note Amount', 'Debit Note Amount']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df['Dealer_Code'] = df['Dealer Location'].map(dealer_mapping).fillna(df['Dealer Location'])
        df['Month'] = df['Fiscal Month'].astype(str).str.strip().str[:3]
        df['Claim arbitration ID'] = df['Claim arbitration ID'].astype(str).replace('nan', '').replace('', np.nan)

        dealers = sorted(df['Dealer_Code'].unique())
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        # CREDIT NOTE
        credit_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Credit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Credit Note {month}']
                credit_df = credit_df.merge(summary, on='Division', how='left')
            else:
                credit_df[f'Credit Note {month}'] = 0
        
        credit_df = credit_df.fillna(0)
        credit_columns = [f'Credit Note {month}' for month in months]
        credit_df['Total Credit'] = credit_df[credit_columns].sum(axis=1)
        grand_total_credit = {'Division': 'Grand Total'}
        for col in credit_df.columns[1:]:
            grand_total_credit[col] = credit_df[col].sum()
        credit_df = pd.concat([credit_df, pd.DataFrame([grand_total_credit])], ignore_index=True)

        # DEBIT NOTE
        debit_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Debit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Debit Note {month}']
                debit_df = debit_df.merge(summary, on='Division', how='left')
            else:
                debit_df[f'Debit Note {month}'] = 0
        
        debit_df = debit_df.fillna(0)
        debit_columns = [f'Debit Note {month}' for month in months]
        debit_df['Total Debit'] = debit_df[debit_columns].sum(axis=1)
        grand_total_debit = {'Division': 'Grand Total'}
        for col in debit_df.columns[1:]:
            grand_total_debit[col] = debit_df[col].sum()
        debit_df = pd.concat([debit_df, pd.DataFrame([grand_total_debit])], ignore_index=True)

        # ARBITRATION
        arbitration_df = pd.DataFrame({'Division': dealers})
        
        def is_arbitration(value):
            if pd.isna(value): return False
            value = str(value).strip().upper()
            return value.startswith('ARB') and value != 'NAN' and value != ''

        for month in months:
            month_data = df[df['Month'] == month].copy()
            arb_records = month_data[month_data['Claim arbitration ID'].apply(is_arbitration)]
            arb_summary = arb_records.groupby('Dealer_Code')['Debit Note Amount'].sum().reset_index()
            arb_summary.columns = ['Division', f'Claim Arbitration {month}']
            arbitration_df = arbitration_df.merge(arb_summary, on='Division', how='left')
        
        arbitration_df = arbitration_df.fillna(0)
        debit_copy = debit_df[debit_df['Division'] != 'Grand Total'][['Division', 'Total Debit']].copy()
        arbitration_df = arbitration_df.merge(debit_copy, on='Division', how='left')
        
        arb_cols = [f'Claim Arbitration {m}' for m in months]
        arbitration_df['Pending Claim Arbitration'] = (
            arbitration_df['Total Debit'] - arbitration_df[arb_cols].sum(axis=1)
        ).apply(lambda x: max(0, x))
        
        arbitration_df = arbitration_df.drop(['Total Debit'], axis=1)
        grand_total_arb = {'Division': 'Grand Total'}
        for col in arbitration_df.columns[1:]:
            grand_total_arb[col] = arbitration_df[col].sum()
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([grand_total_arb])], ignore_index=True)

        print("✓ Warranty Data Processing Complete\n")
        return credit_df, debit_df, arbitration_df, df

    except Exception as e:
        print(f"Error: {e}\n")
        traceback.print_exc()
        return None, None, None, None

def process_current_month_warranty():
    input_path = find_data_file('Pending Warranty Claim Details.xlsx')
    if input_path is None:
        return None, None
    try:
        df = pd.read_excel(input_path, sheet_name='Pending Warranty Claim Details')
        print(f"Processing Current Month Warranty - {len(df)} rows")

        required_columns = ['Division', 'Pending Claims Spares', 'Pending Claims Labour']
        if not all(col in df.columns for col in required_columns):
            return None, None

        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]

        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            summary_data.append({
                'Division': division,
                'Pending Claims Spares Count': int(div_data['Pending Claims Spares'].notna().sum()),
                'Pending Claims Labour Count': int(div_data['Pending Claims Labour'].notna().sum()),
                'Total Pending Claims': int(div_data['Pending Claims Spares'].notna().sum() + div_data['Pending Claims Labour'].notna().sum())
            })
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            'Division': 'Grand Total',
            'Pending Claims Spares Count': int(summary_df['Pending Claims Spares Count'].sum()),
            'Pending Claims Labour Count': int(summary_df['Pending Claims Labour Count'].sum()),
            'Total Pending Claims': int(summary_df['Total Pending Claims'].sum())
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

        print("✓ Current Month Warranty Processing Complete\n")
        return summary_df, df

    except Exception as e:
        print(f"Error: {e}\n")
        traceback.print_exc()
        return None, None

def process_compensation_claim():
    input_path = find_data_file('Transit_Claims_Merged.xlsx')
    if input_path is None:
        return None, None
    try:
        df = pd.read_excel(input_path)
        print(f"Processing Compensation Claims - {len(df)} rows")

        required_columns = ['Division', 'RO Id.', 'Claim Amount', 'Claim Approved Amt.']
        available_columns = [col for col in required_columns if col in df.columns]
        if not available_columns:
            return None, None

        df_filtered = df[available_columns].copy()
        
        if 'Division' in df_filtered.columns:
            df_filtered['Division'] = df_filtered['Division'].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered['Division'].notna() & (df_filtered['Division'] != '') & (df_filtered['Division'] != 'nan')]
        
        numeric_cols = ['Claim Amount', 'Claim Approved Amt.']
        for col in numeric_cols:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
        
        summary_data = []
        if 'Division' in df_filtered.columns:
            for division in sorted(df_filtered['Division'].unique()):
                div_data = df_filtered[df_filtered['Division'] == division]
                summary_row = {'Division': division, 'Total Claims': len(div_data)}
                
                if 'Claim Amount' in df_filtered.columns:
                    summary_row['Total Claim Amount'] = div_data['Claim Amount'].sum()
                if 'Claim Approved Amt.' in df_filtered.columns:
                    summary_row['Total Approved Amount'] = div_data['Claim Approved Amt.'].sum()
                
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            grand_total = {'Division': 'Grand Total'}
            for col in summary_df.columns[1:]:
                if summary_df[col].dtype in ['int64', 'float64']:
                    grand_total[col] = summary_df[col].sum()
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        print("✓ Compensation Claim Processing Complete\n")
        return summary_df, df_filtered

    except Exception as e:
        print(f"Error: {e}\n")
        traceback.print_exc()
        return None, None

def process_pr_approval():
    input_path = find_data_file('Pr_Approval_Claims_Merged.xlsx')
    if input_path is None:
        return None, None
    try:
        df = pd.read_excel(input_path)
        print(f"Processing PR Approval - {len(df)} rows")

        if 'Division' not in df.columns:
            return None, None

        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]

        numeric_cols = ['Total Cost of Repair', 'Req. Claim Amt from M&M', 'App. Claim Amt from M&M']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            summary_row = {'Division': division, 'Total Requests': len(div_data)}
            
            if 'Total Cost of Repair' in df.columns:
                summary_row['Total Cost of Repair'] = div_data['Total Cost of Repair'].sum()
            
            if 'Req. Claim Amt from M&M' in df.columns:
                summary_row['Req. Claim Amt from M&M'] = div_data['Req. Claim Amt from M&M'].sum()
            
            if 'App. Claim Amt from M&M' in df.columns:
                summary_row['Total Approved Amount'] = div_data['App. Claim Amt from M&M'].sum()
            
            summary_data.append(summary_row)
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {'Division': 'Grand Total'}
        for col in summary_df.columns[1:]:
            if summary_df[col].dtype in ['int64', 'float64']:
                grand_total[col] = summary_df[col].sum()
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

        print("✓ PR Approval Processing Complete\n")
        return summary_df, df

    except Exception as e:
        print(f"Error: {e}\n")
        traceback.print_exc()
        return None, None

def load_user_credentials():
    try:
        user_file = os.path.join(DATA_DIR, "UserID.xlsx")
        if not os.path.exists(user_file):
            return {}
        df = pd.read_excel(user_file)
        credentials = {}
        for idx, row in df.iterrows():
            try:
                if 'User ID' in df.columns and 'Password' in df.columns:
                    uid = row['User ID']
                    pwd = row['Password']
                    if pd.notna(uid) and pd.notna(pwd):
                        user_id = str(int(float(uid)))
                        password = str(pwd).strip()
                        credentials[user_id] = password
            except:
                continue
        print(f"✓ Loaded {len(credentials)} user credentials\n")
        return credentials
    except:
        return {}

USER_CREDENTIALS = load_user_credentials()
SESSIONS = {}

class CaptchaGenerator:
    @staticmethod
    def generate_captcha(length=6):
        allowed_chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ123456789'
        captcha_text = ''.join(secrets.choice(allowed_chars) for _ in range(length))
        
        width, height = 500, 150
        image = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(image)
        
        for _ in range(5):
            x1, y1 = secrets.randbelow(width), secrets.randbelow(height)
            x2, y2 = secrets.randbelow(width), secrets.randbelow(height)
            draw.line((x1, y1, x2, y2), fill='lightgray', width=1)
        
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 80)
        except:
            try:
                font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 80)
            except:
                font = ImageFont.load_default()
        
        for i, char in enumerate(captcha_text):
            y_offset = np.random.randint(15, 50)
            draw.text((15 + i * 70, y_offset), char, fill='#FF8C00', font=font)
        
        for _ in range(50):
            x, y = secrets.randbelow(width), secrets.randbelow(height)
            draw.point((x, y), fill='#FFD699')
        
        img_io = io.BytesIO()
        image.save(img_io, 'PNG')
        img_io.seek(0)
        img_base64 = base64.b64encode(img_io.getvalue()).decode()
        return captcha_text, f"data:image/png;base64,{img_base64}"

def create_session(user_id):
    session_id = secrets.token_hex(16)
    SESSIONS[session_id] = {'user_id': user_id, 'created_at': datetime.now(), 'last_activity': datetime.now()}
    return session_id

def verify_session(session_id):
    if session_id not in SESSIONS:
        return None
    session = SESSIONS[session_id]
    if (datetime.now() - session['last_activity']).total_seconds() > 8 * 3600:
        del SESSIONS[session_id]
        return None
    session['last_activity'] = datetime.now()
    return session['user_id']

LOGIN_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<title>Warranty Management - Login</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { width: 100%; height: 100%; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
body { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); min-height: 100vh; display: flex; justify-content: center; align-items: center; padding: 10px; }
.login-wrapper { background: white; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); width: 100%; max-width: 1000px; display: grid; grid-template-columns: 1fr 1fr; max-height: 90vh; overflow-y: auto; }
.login-left { padding: 25px; display: flex; flex-direction: column; justify-content: center; }
.login-right { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); padding: 25px; display: flex; flex-direction: column; justify-content: center; align-items: center; color: white; text-align: center; }
.logo-section h1 { font-size: 22px; color: #333; margin-bottom: 8px; }
.logo-section p { color: #666; font-size: 13px; margin-bottom: 5px; }
.login-form { display: flex; flex-direction: column; gap: 15px; margin-top: 15px; }
.form-group { display: flex; flex-direction: column; gap: 6px; }
.form-group label { font-weight: 600; color: #333; font-size: 12px; }
.form-group input { padding: 10px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px; transition: all 0.3s; }
.form-group input:focus { outline: none; border-color: #FF8C00; box-shadow: 0 0 8px rgba(255,140,0,0.2); }
.captcha-section { margin-top: 12px; padding: 10px; background: #f5f5f5; border-radius: 6px; }
.captcha-image { width: 100%; height: auto; margin-bottom: 8px; border-radius: 4px; }
.login-btn { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; border: none; padding: 11px; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.3s; font-size: 14px; }
.login-btn:active { transform: translateY(1px); }
.error-message { color: #c62828; font-size: 12px; display: none; padding: 8px; background: #ffebee; border-radius: 4px; border-left: 3px solid #c62828; }
.error-message.show { display: block; }
.right-content h2 { font-size: 28px; margin-bottom: 15px; }
.right-content p { font-size: 14px; line-height: 1.5; }

@media (max-width: 768px) {
  .login-wrapper { grid-template-columns: 1fr; }
  .login-left { padding: 20px; }
  .login-right { padding: 20px; min-height: 200px; }
  .logo-section h1 { font-size: 18px; }
  .logo-section p { font-size: 12px; }
  .right-content h2 { font-size: 22px; }
}

@media (max-width: 480px) {
  body { padding: 5px; }
  .login-wrapper { border-radius: 12px; }
  .login-left { padding: 15px; }
  .logo-section h1 { font-size: 16px; }
  .form-group label { font-size: 11px; }
  .form-group input { padding: 9px; font-size: 13px; }
  .login-btn { padding: 10px; font-size: 13px; }
  .error-message { font-size: 11px; }
}
</style>
</head>
<body>
<div class="login-wrapper">
<div class="login-left">
<div class="logo-section">
<h1>Warranty Management</h1>
<p>Enter credentials to access dashboard</p>
</div>
<form class="login-form" onsubmit="handleLogin(event)">
<div class="form-group">
<label for="userId">User ID</label>
<input type="text" id="userId" placeholder="Enter User ID" required autocomplete="username">
</div>
<div class="form-group">
<label for="password">Password</label>
<input type="password" id="password" placeholder="Enter password" required autocomplete="current-password">
</div>
<div class="captcha-section">
<img id="captchaImage" class="captcha-image" src="" alt="CAPTCHA">
<input type="text" id="captchaInput" placeholder="Enter CAPTCHA" required style="width:100%;padding:8px;border:2px solid #e0e0e0;border-radius:4px;font-size:13px;">
</div>
<div class="error-message" id="errorMessage"></div>
<button type="submit" class="login-btn">Login</button>
</form>
</div>
<div class="login-right">
<div class="right-content">
<h2>Welcome</h2>
<p>Warranty Management System</p>
<p style="margin-top:20px;font-size:12px;opacity:0.8;">Unnati Motors</p>
</div>
</div>
</div>
<script>
let currentCaptcha='';
async function loadCaptcha(){const response=await fetch('/api/captcha');const data=await response.json();currentCaptcha=data.captcha;document.getElementById('captchaImage').src=data.image;}
async function handleLogin(event){event.preventDefault();const userId=document.getElementById('userId').value;const password=document.getElementById('password').value;const captchaInput=document.getElementById('captchaInput').value;const errorDiv=document.getElementById('errorMessage');
if(captchaInput.toUpperCase()!==currentCaptcha){errorDiv.textContent='CAPTCHA incorrect';errorDiv.classList.add('show');loadCaptcha();return;}
try{const response=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({user_id:userId,password:password})});
if(response.ok){window.location.href='/dashboard';}else{const error=await response.json();errorDiv.textContent=error.detail||'Login failed';errorDiv.classList.add('show');loadCaptcha();}}catch(error){errorDiv.textContent='Error: '+error.message;errorDiv.classList.add('show');}}
window.onload=loadCaptcha;
</script>
</body>
</html>"""

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes, shrink-to-fit=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="theme-color" content="#FF8C00">
<title>Warranty Dashboard</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { width: 100%; height: 100%; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%); }
.navbar { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 12px 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); position: sticky; top: 0; z-index: 100; }
.navbar h1 { font-size: 18px; font-weight: 700; text-align: center; }
.container { max-width: 1400px; margin: 15px auto; padding: 0 10px; }
.dashboard-content { background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 15px; }
.nav-tabs { border-bottom: 2px solid #FF8C00; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 5px; overflow-x: auto; }
.nav-tabs button { color: #666; font-weight: 600; border: none; border-bottom: 3px solid transparent; padding: 10px 12px; cursor: pointer; transition: all 0.3s; background: none; font-size: 12px; white-space: nowrap; }
.nav-tabs button:hover { color: #FF8C00; }
.nav-tabs button.active { color: #FF8C00; border-bottom-color: #FF8C00; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.export-section { margin: 20px 0; padding: 15px; background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%); border-radius: 8px; border-left: 4px solid #FF8C00; }
.export-section h3 { color: #FF8C00; margin-bottom: 12px; font-size: 14px; }
.export-controls { display: flex; gap: 10px; flex-wrap: wrap; background: white; padding: 12px; border-radius: 6px; }
.export-control-group { display: flex; gap: 5px; align-items: center; min-width: 150px; }
.export-control-group label { font-weight: 600; color: #333; font-size: 12px; min-width: 70px; }
.export-control-group select { padding: 6px; border: 2px solid #FF8C00; border-radius: 4px; cursor: pointer; background: white; font-size: 12px; flex: 1; min-width: 100px; }
.export-btn { padding: 8px 15px; background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; transition: all 0.3s; }
.export-btn:active { transform: translateY(1px); }
.data-table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11px; overflow-x: auto; }
.data-table thead th { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 8px 6px; text-align: center; font-weight: 600; font-size: 10px; border: 1px solid #e0e0e0; }
.data-table tbody td { padding: 7px 5px; border-bottom: 1px solid #e0e0e0; text-align: right; border: 1px solid #e0e0e0; }
.data-table tbody td:first-child { text-align: left; font-weight: 600; color: #333; }
.data-table tbody tr:last-child { background: #fff8f3; font-weight: 700; border-top: 2px solid #FF8C00; color: #FF8C00; }
.table-title { font-size: 13px; font-weight: 700; color: #FF8C00; margin-bottom: 10px; }
.loading-spinner { display: none; text-align: center; padding: 30px; }
.spinner { border: 3px solid rgba(255,140,0,0.2); border-top: 3px solid #FF8C00; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto; }
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
.table-wrapper { overflow-x: auto; border-radius: 6px; }

@media (max-width: 768px) {
  .navbar h1 { font-size: 14px; }
  .container { margin: 10px auto; padding: 0 8px; }
  .dashboard-content { padding: 12px; border-radius: 8px; }
  .nav-tabs { margin-bottom: 15px; gap: 3px; }
  .nav-tabs button { padding: 8px 10px; font-size: 11px; }
  .export-controls { flex-direction: column; gap: 8px; padding: 10px; }
  .export-control-group { flex-direction: column; min-width: auto; }
  .export-control-group label { min-width: auto; margin-bottom: 3px; }
  .export-control-group select { min-width: auto; }
  .export-btn { width: 100%; padding: 10px; font-size: 11px; }
  .data-table { font-size: 10px; }
  .data-table thead th { padding: 6px 4px; font-size: 9px; }
  .data-table tbody td { padding: 5px 4px; }
  .table-title { font-size: 12px; }
}

@media (max-width: 480px) {
  .navbar { padding: 10px 10px; }
  .navbar h1 { font-size: 13px; }
  .container { margin: 8px auto; padding: 0 5px; }
  .dashboard-content { padding: 10px; border-radius: 6px; }
  .nav-tabs { margin-bottom: 12px; gap: 2px; }
  .nav-tabs button { padding: 6px 8px; font-size: 10px; }
  .export-section { margin: 12px 0; padding: 10px; border-left: 3px solid #FF8C00; }
  .export-section h3 { margin-bottom: 8px; font-size: 12px; }
  .export-controls { flex-direction: column; gap: 6px; padding: 8px; }
  .export-control-group { flex-direction: column; }
  .export-control-group label { font-size: 11px; margin-bottom: 2px; }
  .export-control-group select { padding: 5px; font-size: 11px; }
  .export-btn { width: 100%; padding: 9px; font-size: 11px; }
  .data-table { font-size: 9px; margin-top: 10px; }
  .data-table thead th { padding: 5px 3px; font-size: 8px; }
  .data-table tbody td { padding: 4px 3px; }
  .table-title { font-size: 11px; margin-bottom: 8px; }
}
</style>
</head>
<body>
<nav class="navbar">
<h1>Warranty Management Dashboard</h1>
</nav>
<div class="container">
<div class="dashboard-content">
<div class="loading-spinner" id="loadingSpinner">
<div class="spinner"></div>
<p style="margin-top:10px;color:#666;font-size:12px;">Loading data...</p>
</div>
<div id="warrantyTabs" style="display:none;">
<div class="nav-tabs">
<button class="nav-link active" onclick="switchTab('credit')">💳 Credit</button>
<button class="nav-link" onclick="switchTab('debit')">💸 Debit</button>
<button class="nav-link" onclick="switchTab('arbitration')">⚖️ Arbitration</button>
<button class="nav-link" onclick="switchTab('currentmonth')">📅 Current</button>
<button class="nav-link" onclick="switchTab('compensation')">🚗 Compensation</button>
<button class="nav-link" onclick="switchTab('pr_approval')">✅ PR</button>
</div>
<div class="export-section">
<h3>📊 Export to Excel</h3>
<div class="export-controls">
<div class="export-control-group">
<label>Division:</label>
<select id="divisionFilter">
<option value="">-- Select --</option>
<option value="All">All</option>
</select>
</div>
<div class="export-control-group">
<label>Type:</label>
<select id="exportType">
<option value="credit">Credit</option>
<option value="debit">Debit</option>
<option value="arbitration">Arbitration</option>
<option value="currentmonth">Current</option>
<option value="compensation">Compensation</option>
<option value="pr_approval">PR</option>
</select>
</div>
<button onclick="exportToExcel()" class="export-btn">📥 Export</button>
</div>
</div>
<div id="credit" class="tab-content active">
<div class="table-title">Warranty Credit Note</div>
<div class="table-wrapper">
<table class="data-table" id="creditTable"><thead></thead><tbody></tbody></table>
</div>
</div>
<div id="debit" class="tab-content">
<div class="table-title">Warranty Debit Note</div>
<div class="table-wrapper">
<table class="data-table" id="debitTable"><thead></thead><tbody></tbody></table>
</div>
</div>
<div id="arbitration" class="tab-content">
<div class="table-title">Claim Arbitration</div>
<div class="table-wrapper">
<table class="data-table" id="arbitrationTable"><thead></thead><tbody></tbody></table>
</div>
</div>
<div id="currentmonth" class="tab-content">
<div class="table-title">Current Month - Pending</div>
<div class="table-wrapper">
<table class="data-table" id="currentMonthTable"><thead></thead><tbody></tbody></table>
</div>
</div>
<div id="compensation" class="tab-content">
<div class="table-title">Compensation Claim</div>
<div class="table-wrapper">
<table class="data-table" id="compensationTable"><thead></thead><tbody></tbody></table>
</div>
</div>
<div id="pr_approval" class="tab-content">
<div class="table-title">PR Approval</div>
<div class="table-wrapper">
<table class="data-table" id="prApprovalTable"><thead></thead><tbody></tbody></table>
</div>
</div>
</div>
</div>
</div>
<script>
let warrantyData={};
async function loadDashboard(){const spinner=document.getElementById('loadingSpinner');const tabs=document.getElementById('warrantyTabs');spinner.style.display='block';tabs.style.display='none';
try{const response=await fetch('/api/warranty-data',{credentials:'include',headers:{'Content-Type':'application/json'}});
if(response.status===401){window.location.href='/login-page';return;}
if(!response.ok)throw new Error('Failed to load data');
warrantyData=await response.json();populateTable(warrantyData.credit,'creditTable');populateTable(warrantyData.debit,'debitTable');populateTable(warrantyData.arbitration,'arbitrationTable');
populateTable(warrantyData.currentMonth,'currentMonthTable');populateTable(warrantyData.compensation,'compensationTable');populateTable(warrantyData.prApproval,'prApprovalTable');loadDivisions();spinner.style.display='none';tabs.style.display='block';}
catch(error){spinner.innerHTML='<p style="color:red;padding:15px;">Error loading data</p>';}
}
function populateTable(data,tableId){if(!data||data.length===0)return;const table=document.getElementById(tableId);const headers=Object.keys(data[0]);
table.querySelector('thead').innerHTML=headers.map(h=>`<th>${h}</th>`).join('');table.querySelector('tbody').innerHTML=data.map((row)=>
`<tr>${headers.map((h)=>`<td>${typeof row[h]==='number'?row[h].toLocaleString('en-IN',{maximumFractionDigits:0}):row[h]}</td>`).join('')}</tr>`).join('');}
function switchTab(tabName){document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
document.querySelectorAll('.nav-link').forEach(b=>b.classList.remove('active'));document.getElementById(tabName).classList.add('active');event.target.classList.add('active');}
function loadDivisions(){const divisions=new Set();const type=document.getElementById('exportType').value||'credit';
const data=warrantyData[type==='currentmonth'?'currentMonth':type==='compensation'?'compensation':type==='pr_approval'?'prApproval':type];
if(data){data.forEach(row=>{if(row.Division&&row.Division!=='Grand Total'){divisions.add(row.Division);}});}
const select=document.getElementById('divisionFilter');select.innerHTML='<option value="">-- Select --</option><option value="All">All</option>';
Array.from(divisions).sort().forEach(div=>{const opt=document.createElement('option');opt.value=div;opt.textContent=div;select.appendChild(opt);});}
document.getElementById('exportType')?.addEventListener('change',loadDivisions);
async function exportToExcel(){const division=document.getElementById('divisionFilter').value;const type=document.getElementById('exportType').value;if(!division){alert('Select a division');return;}
const btn=document.getElementById('exportBtn');btn.disabled=true;btn.textContent='⏳...';
try{const response=await fetch('/api/export-to-excel',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'include',body:JSON.stringify({division,type})});
if(!response.ok)throw new Error('Export failed');const blob=await response.blob();const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;
a.download=`${type}_${division}_${new Date().toISOString().split('T')[0]}.xlsx`;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}
catch(error){alert('Export failed');}finally{btn.disabled=false;btn.textContent='📥 Export';}}
window.onload=loadDashboard;
</script>
</body>
</html>"""

app = FastAPI()

@app.get("/api/captcha")
async def get_captcha():
    captcha_text, captcha_image = CaptchaGenerator.generate_captcha()
    return {"captcha": captcha_text, "image": captcha_image}

@app.post("/api/login")
async def api_login(request: Request):
    try:
        body = await request.json()
        user_id = body.get('user_id', '').strip()
        password = body.get('password', '')
        
        if not user_id or user_id not in USER_CREDENTIALS:
            raise HTTPException(status_code=401, detail="Invalid User ID")
        
        if USER_CREDENTIALS[user_id] != password:
            raise HTTPException(status_code=401, detail="Invalid Password")
        
        session_id = create_session(user_id)
        response = JSONResponse({"success": True, "session_id": session_id, "user_id": user_id}, status_code=200)
        response.set_cookie(key="session_id", value=session_id, httponly=True, max_age=28800, samesite="lax", path="/")
        return response
    except HTTPException as e:
        raise
    except:
        raise HTTPException(status_code=400, detail="Error")

@app.get("/api/warranty-data")
async def get_warranty_data():
    try:
        if WARRANTY_DATA['credit_df'] is None:
            return {"credit": [], "debit": [], "arbitration": [], "currentMonth": [], "compensation": [], "prApproval": []}
        
        credit_records = WARRANTY_DATA['credit_df'].to_dict('records')
        debit_records = WARRANTY_DATA['debit_df'].to_dict('records')
        arbitration_records = WARRANTY_DATA['arbitration_df'].to_dict('records')
        current_month_records = (WARRANTY_DATA['current_month_df'].to_dict('records') if WARRANTY_DATA['current_month_df'] is not None else [])
        compensation_records = (WARRANTY_DATA['compensation_df'].to_dict('records') if WARRANTY_DATA['compensation_df'] is not None else [])
        pr_approval_records = (WARRANTY_DATA['pr_approval_df'].to_dict('records') if WARRANTY_DATA['pr_approval_df'] is not None else [])
        
        for records in [credit_records, debit_records, arbitration_records, current_month_records, compensation_records, pr_approval_records]:
            for record in records:
                for key in record:
                    if pd.isna(record[key]):
                        record[key] = 0
        
        return {"credit": credit_records, "debit": debit_records, "arbitration": arbitration_records, 
                "currentMonth": current_month_records, "compensation": compensation_records, "prApproval": pr_approval_records}
    except:
        raise HTTPException(status_code=500, detail="Error")

@app.post("/api/export-to-excel")
async def export_to_excel(request: Request):
    try:
        body = await request.json()
        selected_division = body.get('division', 'All')
        export_type = body.get('type', 'credit')
        
        if export_type == 'currentmonth':
            df = WARRANTY_DATA['current_month_df']
        elif export_type == 'compensation':
            df = WARRANTY_DATA['compensation_df']
        elif export_type == 'pr_approval':
            df = WARRANTY_DATA['pr_approval_df']
        elif export_type == 'credit':
            df = WARRANTY_DATA['credit_df']
        elif export_type == 'debit':
            df = WARRANTY_DATA['debit_df']
        else:
            df = WARRANTY_DATA['arbitration_df']
        
        if df is None or df.empty:
            raise HTTPException(status_code=500, detail="No data")
        
        if selected_division != 'All' and selected_division != 'Grand Total':
            df_export = df[df['Division'] == selected_division].copy()
            grand_total_row = df[df['Division'] == 'Grand Total']
            if not grand_total_row.empty:
                df_export = pd.concat([df_export, grand_total_row], ignore_index=True)
        else:
            df_export = df.copy()
        
        wb = Workbook()
        ws = wb.active
        ws.title = export_type[:20]
        
        header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for col_idx, column in enumerate(df_export.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=column)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if isinstance(value, (int, float)):
                    cell.value = value
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.value = str(value) if not pd.isna(value) else ''
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.border = border
        
        for col_idx, column in enumerate(df_export.columns, 1):
            max_length = min(max(df_export[column].astype(str).map(len).max(), len(str(column))) + 2, 30)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_length
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(iter([output.getvalue()]), 
                               media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               headers={"Content-Disposition": f"attachment; filename={filename}"})
    except:
        raise HTTPException(status_code=500, detail="Error")

@app.get("/login-page")
async def login_page():
    return HTMLResponse(content=LOGIN_PAGE)

@app.get("/dashboard")
async def dashboard():
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/")
async def root():
    return HTMLResponse(content=DASHBOARD_HTML)

print("Processing warranty data...\n")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'], WARRANTY_DATA['source_df'] = process_warranty_data()
WARRANTY_DATA['current_month_df'], WARRANTY_DATA['current_month_source_df'] = process_current_month_warranty()
WARRANTY_DATA['compensation_df'], WARRANTY_DATA['compensation_source_df'] = process_compensation_claim()
WARRANTY_DATA['pr_approval_df'], WARRANTY_DATA['pr_approval_source_df'] = process_pr_approval()

if __name__ == "__main__":
    port = int(os.getenv('PORT', 8001))
    print("\n" + "="*100)
    print("✅ SERVER READY - WARRANTY MANAGEMENT SYSTEM (MOBILE RESPONSIVE)")
    print("="*100)
    print(f"Port: {port}")
    print(f"URL: http://localhost:{port}")
    print(f"Mobile URL: http://localhost:{port}")
    print(f"\nTest Credentials:")
    print(f"  User ID: 11724")
    print(f"  Password: un001@123")
    print("="*100 + "\n")
    
    uvicorn.run(app, host="0.0.0.0", port=port)
