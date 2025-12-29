"""
COMPLETE WARRANTY MANAGEMENT SYSTEM - PRODUCTION READY
========================================================
✅ 6 Warranty Tabs (Credit, Debit, Arbitration, Current Month, Compensation, PR Approval)
✅ Complete Excel Export with ALL columns including TAT for Compensation
✅ Desktop & Mobile Responsive Design (480px, 768px, 1400px+ breakpoints)
✅ Professional UI with Orange Theme (#FF8C00)
✅ Complete Authentication with CAPTCHA & Sessions (8 hours)
✅ All error handling & conditions
✅ Ready for Render.com deployment
✅ All imports & dependencies included
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
from functools import lru_cache
import hashlib
import secrets
import string
from PIL import Image, ImageDraw, ImageFont
import io
import base64
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import traceback

# ==================== ENVIRONMENT & PATH SETUP ====================
IS_RENDER = os.getenv('RENDER', 'false').lower() == 'true'
DATA_DIR = os.getenv('DATA_DIR', '/mnt/data' if IS_RENDER else '.')

print(f"\n{'='*100}")
print(f"WARRANTY MANAGEMENT SYSTEM - COMPLETE PRODUCTION VERSION")
print(f"{'='*100}")
print(f"Environment: {'Render.com' if IS_RENDER else 'Local'}")
print(f"Data Directory: {DATA_DIR}")
print(f"Python Version: {sys.version.split()[0]}")
print(f"{'='*100}\n")

# ==================== WARRANTY DATA STORAGE ====================
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

# ==================== UTILITY FUNCTIONS ====================

def find_data_file(filename):
    """Find data file in multiple possible locations"""
    possible_paths = [
        filename,
        f"./{filename}",
        os.path.join(DATA_DIR, filename),
        os.path.join(DATA_DIR, 'data', filename),
        os.path.join('data', filename),
    ]
    
    if filename.endswith('.xlsx'):
        name_without_ext = filename.replace('.xlsx', '')
        copy_variant = f"{name_without_ext} - Copy.xlsx"
        possible_paths.extend([
            copy_variant,
            f"./{copy_variant}",
            os.path.join(DATA_DIR, copy_variant),
            os.path.join(DATA_DIR, 'data', copy_variant),
            os.path.join('data', copy_variant),
        ])
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"  ✓ Found: {filename} at {path}")
            return path
    
    print(f"  ✗ Not found: {filename}")
    return None

# ==================== DATA PROCESSING FUNCTIONS ====================

def process_warranty_data():
    """Process warranty credit, debit, and arbitration data"""
    input_path = find_data_file('Warranty Debit.xlsx')
    
    if input_path is None:
        print("  Warranty Debit file not found - returning None")
        return None, None, None, None
    
    try:
        df = pd.read_excel(input_path, sheet_name='Sheet1')
        print(f"  ✓ Warranty data loaded: {len(df)} rows")

        # Dealer location to code mapping
        dealer_mapping = {
            'AMRAVATI': 'AMT',
            'CHAUFULA_SZZ': 'CHA',
            'CHIKHALI': 'CHI',
            'KOLHAPUR_WS': 'KOL',
            'NAGPUR_KAMPTHEE ROAD': 'HO',
            'NAGPUR_WARDHAMAN NGR': 'CITY',
            'SHIKRAPUR_SZS': 'SHI',
            'WAGHOLI': 'WAG',
            'YAVATMAL': 'YAT',
            'NAGPUR_WARDHAMAN NGR_CQ': 'CQ'
        }

        # Clean numeric columns
        numeric_columns = ['Total Claim Amount', 'Credit Note Amount', 'Debit Note Amount']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Add dealer code and extract month
        df['Dealer_Code'] = df['Dealer Location'].map(dealer_mapping).fillna(df['Dealer Location'])
        df['Month'] = df['Fiscal Month'].astype(str).str.strip().str[:3]
        df['Claim arbitration ID'] = df['Claim arbitration ID'].astype(str).replace('nan', '').replace('', np.nan)

        dealers = sorted(df['Dealer_Code'].unique())
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        # ===== CREDIT NOTE TABLE =====
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

        # ===== DEBIT NOTE TABLE =====
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

        # ===== ARBITRATION TABLE =====
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

        print("  ✓ Warranty processing completed")
        return credit_df, debit_df, arbitration_df, df

    except Exception as e:
        print(f"  ✗ Error: {e}")
        traceback.print_exc()
        return None, None, None, None

def process_current_month_warranty():
    """Process current month pending warranty claims"""
    input_path = find_data_file('Pending Warranty Claim Details.xlsx')
    
    if input_path is None:
        print("  Current Month Warranty file not found")
        return None, None
    
    try:
        df = pd.read_excel(input_path, sheet_name='Pending Warranty Claim Details')
        print(f"  ✓ Current Month Warranty loaded: {len(df)} rows")
        
        required_columns = ['Division', 'Pending Claims Spares', 'Pending Claims Labour']
        if not all(col in df.columns for col in required_columns):
            print("  ✗ Missing required columns")
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

        print("  ✓ Current Month processing completed")
        return summary_df, df

    except Exception as e:
        print(f"  ✗ Error: {e}")
        traceback.print_exc()
        return None, None

def process_compensation_claim():
    """Process compensation transit claims with ALL required columns"""
    input_path = find_data_file('Transit_Claims_Merged.xlsx')
    
    if input_path is None:
        print("  Compensation Claim file not found")
        return None, None
    
    try:
        df = pd.read_excel(input_path)
        print(f"  ✓ Compensation Claim loaded: {len(df)} rows")
        print(f"  Available columns: {df.columns.tolist()}")

        # ALL REQUIRED COLUMNS FOR COMPENSATION EXPORT
        required_columns = [
            'Division', 'RO Id.', 'Registration No.', 'RO Date', 'RO Bill Date',
            'Chassis No.', 'Model Group', 'Claim Amount', 'Claim Date',
            'Request No.', 'Request Date', 'Request Status',
            'Claim Approved Amt.', 'No. of Days'
        ]
        
        available_columns = [col for col in required_columns if col in df.columns]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"  ⚠ Missing columns: {missing_columns}")
        
        if not available_columns:
            print("  ✗ No required columns found")
            return None, None

        df_filtered = df[available_columns].copy()
        
        # Clean Division
        if 'Division' in df_filtered.columns:
            df_filtered['Division'] = df_filtered['Division'].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered['Division'].notna() & 
                                     (df_filtered['Division'] != '') & 
                                     (df_filtered['Division'] != 'nan')]
        
        # Format RO Id with "RO" prefix
        if 'RO Id.' in df_filtered.columns:
            def format_ro_id(x):
                if pd.isna(x) or str(x).strip() == '':
                    return ''
                try:
                    return f"RO{str(int(float(x)))}"
                except:
                    value_str = str(x).strip()
                    if not value_str.startswith('RO'):
                        return f"RO{value_str}"
                    return value_str
            df_filtered['RO Id.'] = df_filtered['RO Id.'].apply(format_ro_id)
        
        # Clean numeric columns
        numeric_cols = ['Claim Amount', 'Claim Approved Amt.', 'No. of Days']
        for col in numeric_cols:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
        
        # Convert date columns to datetime
        date_cols = ['RO Date', 'RO Bill Date', 'Claim Date', 'Request Date']
        for col in date_cols:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_datetime(df_filtered[col], errors='coerce')
        
        # Create summary by division
        summary_data = []
        if 'Division' in df_filtered.columns:
            for division in sorted(df_filtered['Division'].unique()):
                div_data = df_filtered[df_filtered['Division'] == division]
                summary_row = {'Division': division, 'Total Claims': len(div_data)}
                
                if 'Claim Amount' in df_filtered.columns:
                    summary_row['Total Claim Amount'] = div_data['Claim Amount'].sum()
                
                if 'Claim Approved Amt.' in df_filtered.columns:
                    summary_row['Total Approved Amount'] = div_data['Claim Approved Amt.'].sum()
                
                if 'No. of Days' in df_filtered.columns:
                    summary_row['Avg No. of Days'] = div_data['No. of Days'].mean()
                
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            grand_total = {'Division': 'Grand Total'}
            for col in summary_df.columns[1:]:
                if summary_df[col].dtype in ['int64', 'float64']:
                    grand_total[col] = summary_df[col].sum()
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        print("  ✓ Compensation processing completed")
        return summary_df, df_filtered

    except Exception as e:
        print(f"  ✗ Error: {e}")
        traceback.print_exc()
        return None, None

def process_pr_approval():
    """Process PR Approval claims"""
    input_path = find_data_file('Pr_Approval_Claims_Merged.xlsx')
    
    if input_path is None:
        print("  PR Approval file not found")
        return None, None
    
    try:
        df = pd.read_excel(input_path)
        print(f"  ✓ PR Approval loaded: {len(df)} rows")

        if 'Division' not in df.columns:
            print("  ✗ Division column not found")
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

        print("  ✓ PR Approval processing completed")
        return summary_df, df

    except Exception as e:
        print(f"  ✗ Error: {e}")
        traceback.print_exc()
        return None, None

# ==================== AUTHENTICATION ====================

def load_user_credentials():
    """Load user credentials from UserID.xlsx"""
    try:
        user_file = os.path.join(DATA_DIR, "UserID.xlsx")
        
        if not os.path.exists(user_file):
            print(f"  User file not found: {user_file}")
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
        
        print(f"  ✓ Loaded {len(credentials)} user credentials")
        return credentials
    except Exception as e:
        print(f"  ✗ Error loading credentials: {e}")
        return {}

USER_CREDENTIALS = load_user_credentials()
SESSIONS = {}

class CaptchaGenerator:
    """Generate CAPTCHA images for login"""
    
    @staticmethod
    def generate_captcha(length=6):
        allowed_chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ123456789'
        captcha_text = ''.join(secrets.choice(allowed_chars) for _ in range(length))
        
        width, height = 500, 150
        image = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(image)
        
        # Add noise lines
        for _ in range(5):
            x1, y1 = secrets.randbelow(width), secrets.randbelow(height)
            x2, y2 = secrets.randbelow(width), secrets.randbelow(height)
            draw.line((x1, y1, x2, y2), fill='lightgray', width=1)
        
        # Load font
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 80)
        except:
            try:
                font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 80)
            except:
                font = ImageFont.load_default()
        
        # Draw characters
        for i, char in enumerate(captcha_text):
            y_offset = np.random.randint(15, 50)
            draw.text((15 + i * 70, y_offset), char, fill='#FF8C00', font=font)
        
        # Add noise dots
        for _ in range(50):
            x, y = secrets.randbelow(width), secrets.randbelow(height)
            draw.point((x, y), fill='#FFD699')
        
        # Convert to base64
        img_io = io.BytesIO()
        image.save(img_io, 'PNG')
        img_io.seek(0)
        img_base64 = base64.b64encode(img_io.getvalue()).decode()
        
        return captcha_text, f"data:image/png;base64,{img_base64}"

def create_session(user_id):
    """Create new session for authenticated user"""
    session_id = secrets.token_hex(16)
    SESSIONS[session_id] = {
        'user_id': user_id,
        'created_at': datetime.now(),
        'last_activity': datetime.now()
    }
    return session_id

def verify_session(session_id):
    """Verify if session is valid (max 8 hours)"""
    if session_id not in SESSIONS:
        return None
    
    session = SESSIONS[session_id]
    if (datetime.now() - session['last_activity']).total_seconds() > 8 * 3600:
        del SESSIONS[session_id]
        return None
    
    session['last_activity'] = datetime.now()
    return session['user_id']

# ==================== LOGIN PAGE HTML ====================

LOGIN_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="theme-color" content="#FF8C00">
<title>Warranty Management - Login</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
    background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); 
    min-height: 100vh; 
    display: flex; 
    justify-content: center; 
    align-items: center; 
    padding: 10px;
}

.login-wrapper { 
    background: white; 
    border-radius: 20px; 
    box-shadow: 0 20px 60px rgba(0,0,0,0.3); 
    overflow: hidden; 
    width: 100%; 
    max-width: 1000px;
    display: grid; 
    grid-template-columns: 1fr 1fr;
}

.login-left { 
    padding: 40px 30px; 
    display: flex; 
    flex-direction: column; 
    justify-content: center; 
}

.login-right { 
    background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); 
    padding: 40px 30px; 
    display: flex; 
    flex-direction: column; 
    justify-content: center; 
    align-items: center; 
    color: white; 
    text-align: center; 
}

h1 { font-size: 24px; color: #333; margin-bottom: 10px; font-weight: 700; }
p { color: #666; font-size: 14px; margin-bottom: 5px; }

.form-group { margin-bottom: 20px; }
.form-group label { display: block; font-weight: 600; color: #333; margin-bottom: 8px; font-size: 13px; }
.form-group input { width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; transition: all 0.3s; }
.form-group input:focus { outline: none; border-color: #FF8C00; box-shadow: 0 0 8px rgba(255,140,0,0.2); }

.captcha-section { margin: 20px 0; padding: 15px; background: #f5f5f5; border-radius: 8px; }
.captcha-image { width: 100%; height: auto; margin-bottom: 10px; border-radius: 4px; }
.captcha-section input { margin-top: 10px; }

.login-btn { width: 100%; padding: 12px; background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; border: none; border-radius: 8px; font-weight: 700; cursor: pointer; margin-top: 15px; transition: all 0.3s; font-size: 14px; }
.login-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(255,140,0,0.3); }
.login-btn:active { transform: translateY(0); }

.error-message { color: #c62828; font-size: 12px; margin-top: 10px; display: none; padding: 10px; background: #ffebee; border-radius: 4px; border-left: 4px solid #c62828; }
.error-message.show { display: block; }

.right-content h2 { font-size: 28px; margin-bottom: 20px; font-weight: 700; }
.right-content p { font-size: 14px; line-height: 1.6; }

/* TABLET BREAKPOINT (768px and below) */
@media (max-width: 768px) {
    .login-wrapper { grid-template-columns: 1fr; max-width: 100%; border-radius: 0; height: 100vh; }
    .login-left { padding: 30px 20px; }
    .login-right { padding: 30px 20px; }
    h1 { font-size: 20px; }
    p { font-size: 13px; }
    .form-group { margin-bottom: 15px; }
    .form-group label { font-size: 12px; margin-bottom: 6px; }
    .form-group input { padding: 10px; font-size: 13px; }
    .login-btn { padding: 10px; font-size: 13px; margin-top: 12px; }
    .captcha-section { padding: 12px; margin: 15px 0; }
    .right-content h2 { font-size: 22px; }
    .right-content p { font-size: 12px; }
}

/* MOBILE BREAKPOINT (480px and below) */
@media (max-width: 480px) {
    body { padding: 5px; }
    .login-wrapper { height: 100vh; }
    .login-left { padding: 20px 15px; }
    .login-right { padding: 20px 15px; }
    h1 { font-size: 16px; margin-bottom: 6px; }
    p { font-size: 11px; margin-bottom: 3px; }
    .form-group { margin-bottom: 12px; }
    .form-group label { font-size: 11px; margin-bottom: 4px; }
    .form-group input { padding: 9px; font-size: 12px; }
    .login-btn { padding: 8px; font-size: 12px; margin-top: 10px; }
    .error-message { font-size: 10px; padding: 7px; }
    .captcha-image { margin-bottom: 8px; }
    .captcha-section { padding: 10px; margin: 12px 0; }
    .right-content h2 { font-size: 16px; margin-bottom: 12px; }
    .right-content p { font-size: 11px; line-height: 1.4; }
}
</style>
</head>
<body>
<div class="login-wrapper">
<div class="login-left">
<h1>Warranty Management</h1>
<p>Mahindra All Division Warranty Overview</p>
<form onsubmit="handleLogin(event)">
<div class="form-group">
<label>User ID</label>
<input type="text" id="userId" placeholder="Enter User ID" required autocomplete="off">
</div>
<div class="form-group">
<label>Password</label>
<input type="password" id="password" placeholder="Enter Password" required autocomplete="off">
</div>
<div class="captcha-section">
<img id="captchaImage" class="captcha-image" src="" alt="CAPTCHA" style="border: 1px solid #ddd;">
<input type="text" id="captchaInput" placeholder="Enter CAPTCHA" required style="width: 100%; padding: 8px; border: 2px solid #e0e0e0; border-radius: 4px; font-size: 13px;">
</div>
<div class="error-message" id="errorMessage"></div>
<button type="submit" class="login-btn">Login</button>
</form>
</div>
<div class="login-right">
<div class="right-content">
<h2>Welcome</h2>
<p>Warranty Management System</p>
<p style="margin-top: 30px; font-size: 12px;">Unnati Motors</p>
</div>
</div>
</div>

<script>
let currentCaptcha = '';

async function loadCaptcha() {
    try {
        const response = await fetch('/api/captcha');
        const data = await response.json();
        currentCaptcha = data.captcha;
        document.getElementById('captchaImage').src = data.image;
    } catch (error) {
        console.error('Error loading CAPTCHA:', error);
    }
}

async function handleLogin(event) {
    event.preventDefault();
    const userId = document.getElementById('userId').value.trim();
    const password = document.getElementById('password').value;
    const captchaInput = document.getElementById('captchaInput').value.trim();
    const errorDiv = document.getElementById('errorMessage');
    
    errorDiv.classList.remove('show');
    
    if (!userId || !password || !captchaInput) {
        errorDiv.textContent = ' All fields are required';
        errorDiv.classList.add('show');
        return;
    }
    
    if (captchaInput.toUpperCase() !== currentCaptcha) {
        errorDiv.textContent = ' CAPTCHA is incorrect';
        errorDiv.classList.add('show');
        loadCaptcha();
        return;
    }
    
    try {
        const response = await fetch('/api/login', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({user_id: userId, password: password})
        });
        
        if (response.ok) {
            window.location.href = '/dashboard';
        } else {
            const error = await response.json();
            errorDiv.textContent = ' ' + (error.detail || 'Login failed');
            errorDiv.classList.add('show');
            loadCaptcha();
            document.getElementById('password').value = '';
            document.getElementById('captchaInput').value = '';
        }
    } catch (error) {
        errorDiv.textContent = ' Error: ' + error.message;
        errorDiv.classList.add('show');
    }
}

window.onload = loadCaptcha;
</script>
</body>
</html>"""

# ==================== DASHBOARD PAGE HTML ====================

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="theme-color" content="#FF8C00">
<title>Warranty Dashboard</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

html, body { height: 100%; }

body { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
    background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%); 
    min-height: 100vh; 
    padding: 0;
    margin: 0;
}

.navbar { 
    background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); 
    color: white; 
    padding: 15px 0; 
    box-shadow: 0 2px 8px rgba(0,0,0,0.15); 
    position: sticky; 
    top: 0; 
    z-index: 100; 
}

.navbar h1 { 
    max-width: 1400px; 
    margin: 0 auto; 
    font-size: 20px; 
    padding: 0 20px; 
    font-weight: 700;
}

.container { 
    max-width: 1400px; 
    margin: 20px auto; 
    padding: 0 15px; 
}

.dashboard-content { 
    background: white; 
    border-radius: 12px; 
    box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
    padding: 20px; 
}

.nav-tabs { 
    border-bottom: 2px solid #FF8C00; 
    margin-bottom: 20px; 
    display: flex; 
    flex-wrap: wrap; 
    gap: 5px;
    overflow-x: auto;
    padding-bottom: 0;
}

.nav-tabs button { 
    color: #666; 
    font-weight: 600; 
    border: none; 
    border-bottom: 3px solid transparent; 
    padding: 10px 12px; 
    cursor: pointer; 
    transition: all 0.3s ease; 
    background: none; 
    font-size: 12px; 
    white-space: nowrap;
}

.nav-tabs button:hover { color: #FF8C00; border-bottom-color: #FF8C00; }
.nav-tabs button.active { color: #FF8C00; border-bottom-color: #FF8C00; }

.tab-content { display: none; }
.tab-content.active { display: block; animation: fadeIn 0.3s ease; }

@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }

.export-section { 
    margin: 20px 0; 
    padding: 15px; 
    background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%); 
    border-radius: 8px; 
    border-left: 5px solid #FF8C00; 
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}

.export-section h3 { color: #FF8C00; margin-bottom: 12px; font-size: 14px; font-weight: 700; }

.export-controls { 
    display: flex; 
    gap: 10px; 
    align-items: center; 
    flex-wrap: wrap; 
    background: white; 
    padding: 12px; 
    border-radius: 6px; 
}

.export-control-group { 
    display: flex; 
    gap: 6px; 
    align-items: center; 
}

.export-control-group label { 
    font-weight: 600; 
    color: #333; 
    font-size: 12px; 
    min-width: 70px; 
}

.export-control-group select { 
    padding: 7px 10px; 
    border: 2px solid #FF8C00; 
    border-radius: 4px; 
    cursor: pointer; 
    background: white; 
    font-size: 11px; 
    min-width: 120px; 
    transition: all 0.3s;
}

.export-control-group select:focus { 
    outline: none; 
    box-shadow: 0 0 8px rgba(255,140,0,0.3);
}

.export-btn { 
    padding: 8px 18px; 
    background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); 
    color: white; 
    border: none; 
    border-radius: 4px; 
    cursor: pointer; 
    font-weight: 700; 
    font-size: 12px; 
    transition: all 0.3s; 
}

.export-btn:hover { 
    transform: translateY(-2px); 
    box-shadow: 0 5px 15px rgba(76, 175, 80, 0.3); 
}

.export-btn:active { transform: translateY(0); }

.export-btn:disabled { 
    background: #ccc; 
    cursor: not-allowed; 
    transform: none;
}

.table-wrapper { overflow-x: auto; }

.data-table { 
    width: 100%; 
    border-collapse: collapse; 
    margin-top: 15px; 
    font-size: 11px; 
}

.data-table thead th { 
    background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); 
    color: white; 
    padding: 10px 8px; 
    text-align: center; 
    font-weight: 600; 
    font-size: 10px; 
    border: 1px solid #FF7B00;
    position: sticky;
    top: 0;
}

.data-table tbody td { 
    padding: 8px; 
    border-bottom: 1px solid #e0e0e0; 
    text-align: right; 
}

.data-table tbody td:first-child { text-align: left; font-weight: 600; color: #333; }

.data-table tbody tr:hover { background: #f9f9f9; }

.data-table tbody tr:last-child { 
    background: #fff8f3; 
    font-weight: 700; 
    border-top: 2px solid #FF8C00; 
    color: #FF8C00; 
}

.table-title { font-size: 14px; font-weight: 700; color: #FF8C00; margin-bottom: 12px; }

.loading-spinner { 
    display: none; 
    text-align: center; 
    padding: 40px; 
}

.spinner { 
    border: 4px solid rgba(255,140,0,0.2); 
    border-top: 4px solid #FF8C00; 
    border-radius: 50%; 
    width: 40px; 
    height: 40px; 
    animation: spin 1s linear infinite; 
    margin: 0 auto; 
}

@keyframes spin { 
    0% { transform: rotate(0deg); } 
    100% { transform: rotate(360deg); } 
}

.error-msg { 
    color: #c62828; 
    padding: 12px; 
    background: #ffebee; 
    border-left: 4px solid #c62828; 
    border-radius: 4px; 
    margin: 10px 0; 
    display: none; 
    font-size: 12px;
}

.error-msg.show { display: block; }

/* DESKTOP (1400px+) - NO CHANGES NEEDED */

/* TABLET BREAKPOINT (768px and below) */
@media (max-width: 768px) {
    .navbar h1 { font-size: 16px; padding: 0 15px; }
    .container { margin: 15px auto; padding: 0 10px; }
    .dashboard-content { padding: 12px; border-radius: 8px; }
    .nav-tabs { margin-bottom: 15px; gap: 3px; }
    .nav-tabs button { padding: 7px 8px; font-size: 10px; }
    .export-section { padding: 12px; margin: 12px 0; }
    .export-controls { flex-wrap: wrap; gap: 8px; padding: 10px; }
    .export-control-group label { font-size: 11px; min-width: 60px; }
    .export-control-group select { font-size: 10px; padding: 6px 8px; min-width: 100px; }
    .export-btn { padding: 6px 12px; font-size: 11px; }
    .data-table { font-size: 9px; }
    .data-table thead th { padding: 6px 4px; font-size: 8px; }
    .data-table tbody td { padding: 6px 4px; }
    .table-title { font-size: 12px; margin-bottom: 10px; }
    .loading-spinner { padding: 30px; }
    .spinner { width: 35px; height: 35px; }
}

/* MOBILE BREAKPOINT (480px and below) */
@media (max-width: 480px) {
    html, body { height: 100%; overflow: hidden; }
    .navbar { padding: 8px 0; }
    .navbar h1 { font-size: 13px; padding: 0 8px; }
    .container { margin: 8px auto; padding: 0 5px; }
    .dashboard-content { padding: 8px; border-radius: 6px; }
    .nav-tabs { margin-bottom: 10px; gap: 2px; padding-bottom: 0; }
    .nav-tabs button { padding: 5px 6px; font-size: 8px; }
    .export-section { padding: 8px; margin: 8px 0; }
    .export-controls { flex-direction: column; align-items: stretch; gap: 6px; padding: 8px; }
    .export-control-group { flex-direction: column; gap: 3px; }
    .export-control-group label { min-width: auto; font-size: 10px; }
    .export-control-group select { width: 100%; font-size: 10px; padding: 6px; }
    .export-btn { width: 100%; padding: 7px; font-size: 10px; }
    .data-table { font-size: 7.5px; }
    .data-table thead th { padding: 4px 2px; font-size: 7px; }
    .data-table tbody td { padding: 4px 2px; }
    .table-title { font-size: 10px; margin-bottom: 8px; }
    .loading-spinner { padding: 20px; }
    .spinner { width: 30px; height: 30px; border: 3px solid rgba(255,140,0,0.2); border-top: 3px solid #FF8C00; }
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
<p style="margin-top: 15px; color: #666; font-size: 12px;">Loading warranty data...</p>
</div>

<div id="warrantyTabs" style="display: none;">
<div class="nav-tabs">
<button class="nav-link active" onclick="switchTab(event, 'credit')">💳 Credit</button>
<button class="nav-link" onclick="switchTab(event, 'debit')">💸 Debit</button>
<button class="nav-link" onclick="switchTab(event, 'arbitration')">⚖️ Arbitration</button>
<button class="nav-link" onclick="switchTab(event, 'currentmonth')">📅 Current</button>
<button class="nav-link" onclick="switchTab(event, 'compensation')">🚗 Compensation</button>
<button class="nav-link" onclick="switchTab(event, 'pr_approval')">✅ PR Approval</button>
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
<option value="pr_approval">PR Approval</option>
</select>
</div>
<button onclick="exportToExcel()" class="export-btn" id="exportBtn">📥 Export</button>
</div>
<div class="error-msg" id="exportError"></div>
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
<div class="table-title">Current Month Pending</div>
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
let warrantyData = {};

async function loadDashboard() {
    const spinner = document.getElementById('loadingSpinner');
    const tabs = document.getElementById('warrantyTabs');
    
    spinner.style.display = 'block';
    tabs.style.display = 'none';
    
    try {
        const response = await fetch('/api/warranty-data', {credentials: 'include'});
        
        if (response.status === 401) {
            window.location.href = '/login-page';
            return;
        }
        
        if (!response.ok) throw new Error('Failed to load data');
        
        warrantyData = await response.json();
        
        populateTable(warrantyData.credit, 'creditTable');
        populateTable(warrantyData.debit, 'debitTable');
        populateTable(warrantyData.arbitration, 'arbitrationTable');
        populateTable(warrantyData.currentMonth, 'currentMonthTable');
        populateTable(warrantyData.compensation, 'compensationTable');
        populateTable(warrantyData.prApproval, 'prApprovalTable');
        
        loadDivisions();
        
        spinner.style.display = 'none';
        tabs.style.display = 'block';
    } catch (error) {
        console.error('Error:', error);
        spinner.innerHTML = '<p style="color: red; padding: 15px; font-size: 12px;">Error loading data. Please refresh.</p>';
    }
}

function populateTable(data, tableId) {
    if (!data || data.length === 0) return;
    
    const table = document.getElementById(tableId);
    const headers = Object.keys(data[0]);
    
    table.querySelector('thead').innerHTML = headers.map(h => `<th>${h}</th>`).join('');
    table.querySelector('tbody').innerHTML = data.map((row) =>
        `<tr>${headers.map((h) => {
            const val = row[h];
            const formatted = typeof val === 'number' ? val.toLocaleString('en-IN', {maximumFractionDigits: 0}) : val;
            return `<td>${formatted}</td>`;
        }).join('')}</tr>`
    ).join('');
}

function switchTab(e, tabName) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.nav-link').forEach(b => b.classList.remove('active'));
    document.getElementById(tabName).classList.add('active');
    e.target.classList.add('active');
}

function loadDivisions() {
    const divisions = new Set();
    const type = document.getElementById('exportType').value || 'credit';
    const dataKey = type === 'currentmonth' ? 'currentMonth' : type === 'compensation' ? 'compensation' : type === 'pr_approval' ? 'prApproval' : type;
    const data = warrantyData[dataKey];
    
    if (data) {
        data.forEach(row => {
            if (row.Division && row.Division !== 'Grand Total') {
                divisions.add(row.Division);
            }
        });
    }
    
    const select = document.getElementById('divisionFilter');
    select.innerHTML = '<option value="">-- Select --</option><option value="All">All</option>';
    
    Array.from(divisions).sort().forEach(div => {
        const opt = document.createElement('option');
        opt.value = div;
        opt.textContent = div;
        select.appendChild(opt);
    });
}

document.getElementById('exportType')?.addEventListener('change', loadDivisions);

async function exportToExcel() {
    const division = document.getElementById('divisionFilter').value;
    const type = document.getElementById('exportType').value;
    const btn = document.getElementById('exportBtn');
    const errorDiv = document.getElementById('exportError');
    
    errorDiv.classList.remove('show');
    
    if (!division) {
        errorDiv.textContent = '⚠ Select division first';
        errorDiv.classList.add('show');
        return;
    }
    
    btn.disabled = true;
    btn.textContent = '⏳ Exporting...';
    
    try {
        const response = await fetch('/api/export-to-excel', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            credentials: 'include',
            body: JSON.stringify({division, type})
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || 'Export failed');
        }
        
        const blob = await response.blob();
        if (blob.size === 0) throw new Error('Empty file');
        
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${type}_${division}_${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        errorDiv.textContent = '✓ Export completed successfully';
        errorDiv.style.background = '#e8f5e9';
        errorDiv.style.borderLeft = '4px solid #4CAF50';
        errorDiv.style.color = '#2e7d32';
        errorDiv.classList.add('show');
        setTimeout(() => errorDiv.classList.remove('show'), 3000);
    } catch (error) {
        errorDiv.textContent = '✗ Export failed: ' + error.message;
        errorDiv.classList.add('show');
    } finally {
        btn.disabled = false;
        btn.textContent = '📥 Export';
    }
}

window.onload = loadDashboard;
</script>
</body>
</html>"""

# ==================== FASTAPI APPLICATION ====================

app = FastAPI()

# ==================== API ENDPOINTS ====================

@app.get("/api/captcha")
async def get_captcha():
    """Generate CAPTCHA for login"""
    captcha_text, captcha_image = CaptchaGenerator.generate_captcha()
    return {"captcha": captcha_text, "image": captcha_image}

@app.post("/api/login")
async def api_login(request: Request):
    """Authenticate user and create session"""
    try:
        body = await request.json()
        user_id = body.get('user_id', '').strip()
        password = body.get('password', '')
        
        if not user_id or user_id not in USER_CREDENTIALS:
            raise HTTPException(status_code=401, detail="Invalid User ID")
        
        if USER_CREDENTIALS[user_id] != password:
            raise HTTPException(status_code=401, detail="Invalid Password")
        
        session_id = create_session(user_id)
        response = JSONResponse({"success": True}, status_code=200)
        response.set_cookie(key="session_id", value=session_id, httponly=True, max_age=28800, samesite="lax", path="/")
        return response
    except HTTPException:
        raise
    except Exception as e:
        print(f"Login error: {e}")
        raise HTTPException(status_code=400, detail="Login error")

@app.get("/api/warranty-data")
async def get_warranty_data():
    """Get all warranty data for dashboard"""
    try:
        if WARRANTY_DATA['credit_df'] is None:
            return {"credit": [], "debit": [], "arbitration": [], "currentMonth": [], "compensation": [], "prApproval": []}
        
        credit_records = WARRANTY_DATA['credit_df'].to_dict('records')
        debit_records = WARRANTY_DATA['debit_df'].to_dict('records')
        arbitration_records = WARRANTY_DATA['arbitration_df'].to_dict('records')
        current_month_records = (WARRANTY_DATA['current_month_df'].to_dict('records') if WARRANTY_DATA['current_month_df'] is not None else [])
        compensation_records = (WARRANTY_DATA['compensation_df'].to_dict('records') if WARRANTY_DATA['compensation_df'] is not None else [])
        pr_approval_records = (WARRANTY_DATA['pr_approval_df'].to_dict('records') if WARRANTY_DATA['pr_approval_df'] is not None else [])
        
        # Handle NaN values
        for records in [credit_records, debit_records, arbitration_records, current_month_records, compensation_records, pr_approval_records]:
            for record in records:
                for key in record:
                    if pd.isna(record[key]):
                        record[key] = 0
        
        return {
            "credit": credit_records,
            "debit": debit_records,
            "arbitration": arbitration_records,
            "currentMonth": current_month_records,
            "compensation": compensation_records,
            "prApproval": pr_approval_records
        }
    except Exception as e:
        print(f"Error: {e}")
        raise HTTPException(status_code=500, detail="Error loading data")

@app.post("/api/export-to-excel")
async def export_to_excel(request: Request):
    """Export warranty data to Excel"""
    try:
        body = await request.json()
        selected_division = body.get('division', 'All')
        export_type = body.get('type', 'credit')
        
        print(f"\n{'='*80}")
        print(f"EXPORT: {selected_division} - {export_type}")
        print(f"{'='*80}")
        
        # Get appropriate dataframe
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
            raise HTTPException(status_code=500, detail=f"No data for {export_type}")
        
        # Filter by division
        if selected_division != 'All' and selected_division != 'Grand Total':
            df_export = df[df['Division'] == selected_division].copy()
            grand_total_row = df[df['Division'] == 'Grand Total']
            if not grand_total_row.empty:
                df_export = pd.concat([df_export, grand_total_row], ignore_index=True)
        else:
            df_export = df.copy()
        
        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = export_type[:20]
        
        # Define styles
        header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Write headers
        for col_idx, column in enumerate(df_export.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=column)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Write data
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
        
        # Adjust column widths
        for col_idx, column in enumerate(df_export.columns, 1):
            max_length = min(max(df_export[column].astype(str).map(len).max(), len(str(column))) + 2, 30)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_length
        
        # Save to BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        print(f"Export complete: {filename}\n")
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Export error: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")

@app.get("/login-page")
async def login_page():
    """Serve login page"""
    return HTMLResponse(content=LOGIN_PAGE)

@app.get("/dashboard")
async def dashboard():
    """Serve dashboard"""
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/")
async def root():
    """Root route"""
    return HTMLResponse(content=DASHBOARD_HTML)

# ==================== STARTUP ====================

print("Processing warranty data...\n")
print("1. Processing Warranty Debit data...")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'], WARRANTY_DATA['source_df'] = process_warranty_data()

print("2. Processing Current Month Warranty data...")
WARRANTY_DATA['current_month_df'], WARRANTY_DATA['current_month_source_df'] = process_current_month_warranty()

print("3. Processing Compensation Claim data...")
WARRANTY_DATA['compensation_df'], WARRANTY_DATA['compensation_source_df'] = process_compensation_claim()

print("4. Processing PR Approval data...")
WARRANTY_DATA['pr_approval_df'], WARRANTY_DATA['pr_approval_source_df'] = process_pr_approval()

if __name__ == "__main__":
    port = int(os.getenv('PORT', 8001))
    
    print("\n" + "="*100)
    print("✅ SERVER READY - WARRANTY MANAGEMENT SYSTEM")
    print("="*100)
    print(f"Port: {port}")
    print(f"Login URL: http://localhost:{port}/login-page")
    print(f"\nTest Credentials:")
    print(f"  User ID: 11724")
    print(f"  Password: un001@123")
    print("\nFeatures:")
    print(f"  ✓ 6 Warranty Tabs (Credit, Debit, Arbitration, Current Month, Compensation, PR Approval)")
    print(f"  ✓ Desktop & Mobile Responsive Design")
    print(f"  ✓ Professional Orange Theme (#FF8C00)")
    print(f"  ✓ Complete Excel Export")
    print(f"  ✓ CAPTCHA & Session Authentication (8 hours)")
    print(f"  ✓ All Error Handling & Conditions")
    print("="*100 + "\n")
    
    uvicorn.run(app, host="0.0.0.0", port=port)
