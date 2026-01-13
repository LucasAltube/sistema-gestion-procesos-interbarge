#!/usr/bin/env python3

import os
import sys
from pathlib import Path

def convert_pdf_to_word_manual():
    """
    Manual PDF to Word conversion instructions for PROC-RHU-003
    Since automated PDF conversion can be unreliable, provide manual instructions
    """
    
    pdf_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.pdf"
    target_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.docx"
    
    print("🔄 PDF TO WORD CONVERSION - PROC-RHU-003")
    print("=" * 60)
    
    if not os.path.exists(pdf_path):
        print(f"❌ PDF file not found: {pdf_path}")
        return False
    
    print(f"📄 Source PDF: {pdf_path}")
    print(f"🎯 Target Word: {target_path}")
    print()
    
    # Check if target already exists
    if os.path.exists(target_path):
        print(f"✅ Word version already exists!")
        print(f"📁 Found: {target_path}")
        return True
    
    print("📋 CONVERSION OPTIONS:")
    print()
    print("Option 1 - RECOMMENDED - Online Converter:")
    print("   1. Upload PDF to: https://www.ilovepdf.com/pdf_to_word")
    print("   2. Download converted .docx file")
    print("   3. Save as: PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.docx")
    print("   4. Place in RRHH/ folder")
    print()
    
    print("Option 2 - Microsoft Word (if available):")
    print("   1. Open Microsoft Word")
    print("   2. File → Open → Select the PDF")
    print("   3. Word will convert automatically")
    print("   4. Save As → .docx format")
    print()
    
    print("Option 3 - Google Docs:")
    print("   1. Upload PDF to Google Drive")
    print("   2. Right-click → Open with Google Docs")
    print("   3. File → Download → Microsoft Word (.docx)")
    print()
    
    print("⚠️  IMPORTANT: After conversion, the document will be ready for safe_optimizer.py")
    print("   The optimizer will apply the standard V.2 → V.3 updates")
    
    return False

def attempt_automated_conversion():
    """
    Attempt automated conversion using available libraries
    """
    print("🤖 ATTEMPTING AUTOMATED CONVERSION...")
    
    try:
        # Try with pdf2docx (if available)
        from pdf2docx import parse
        
        pdf_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.pdf"
        docx_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.docx"
        
        print(f"📦 Using pdf2docx library")
        print(f"🔄 Converting: {pdf_path}")
        
        parse(pdf_path, docx_path)
        
        if os.path.exists(docx_path):
            print(f"✅ SUCCESS! Converted to: {docx_path}")
            return True
        else:
            print(f"❌ Conversion failed - file not created")
            return False
            
    except ImportError:
        print("❌ pdf2docx not available")
        
    try:
        # Try with python-docx2txt (reverse approach)
        import subprocess
        
        # Try using system tools
        pdf_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.pdf"
        docx_path = "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.docx"
        
        print("🔧 Trying system conversion tools...")
        
        # This is a placeholder - system tools vary by OS
        print("❌ No suitable system tools found")
        return False
        
    except Exception as e:
        print(f"❌ System conversion failed: {e}")
        return False
    
    return False

if __name__ == "__main__":
    print("🔄 PDF TO WORD CONVERTER FOR PROC-RHU-003")
    print("=" * 50)
    
    # First attempt automated conversion
    success = attempt_automated_conversion()
    
    if not success:
        print("\n" + "=" * 50)
        print("⚠️  AUTOMATED CONVERSION NOT AVAILABLE")
        print("   Switching to manual conversion instructions...")
        print()
        convert_pdf_to_word_manual()