#!/usr/bin/env python3

import os
import sys
import subprocess
from pathlib import Path

# Complete inventory of all documents to process - UPDATED WITH ALL DOCUMENTS
DOCUMENTS_INVENTORY = [
    # ADMINISTRACION - ALL existing files
    {"code": "PROC-ADM-001", "input": "Administracion/Control de gestion y reporting/PROC-ADM-001 Gestion de Weekly Report V.2.docx", "output": "PROC-ADM-001 Gestión de Weekly Report V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-002", "input": "Administracion/Control de gestion y reporting/PROC-ADM-002 Gestion de Proyecciones.docx", "output": "PROC-ADM-002 Gestión de Proyecciones V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-003", "input": "Administracion/Tesoreria/PROC-ADM-003 Gestion de Pagos V.2.docx", "output": "PROC-ADM-003 Gestión de Pagos V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-004", "input": "Administracion/Tesoreria/PROC-ADM-004 Gestion de Dashboard V.2.docx", "output": "PROC-ADM-004 Gestión de Dashboard V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-005", "input": "Administracion/Tesoreria/PROC-ADM-005 Proyección de Caja V.2.docx", "output": "PROC-ADM-005 Proyección de Caja V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-006", "input": "Administracion/Control de gestion y reporting/PROC-ADM-006 Gestion de Presupuestación.docx", "output": "PROC-ADM-006 Gestión de Presupuestación V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-007", "input": "Administracion/Contabilidad/PROC-ADM-007 Gestion de Cierre Contable V.2.docx", "output": "PROC-ADM-007 Gestión de Cierre Contable V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-008", "input": "Administracion/Contabilidad/PROC-ADM-008 Conciliación de Proveedores 2024.docx", "output": "PROC-ADM-008 Conciliación de Proveedores V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-009", "input": "Administracion/Contabilidad/PROC-ADM-009 Gestión de Facturación 2024.docx", "output": "PROC-ADM-009 Gestión de Facturación V.3.docx", "area": "Administracion"},
    {"code": "PROC-ADM-010", "input": "Administracion/Contabilidad/PROC-ADM-010 Gestión de Cobranzas 2024.docx", "output": "PROC-ADM-010 Gestión de Cobranzas V.3.docx", "area": "Administracion"},
    {"code": "INST-ADM-001", "input": "Administracion/Contabilidad/INST-ADM-001 Instructivo de Facturación en sistema Netsuite.docx", "output": "INST-ADM-001 Instructivo de Facturación en sistema Netsuite V.3.docx", "area": "Administracion"},
    
    # COMERCIAL
    {"code": "INST-COM-001", "input": "Comercial/INST-COM-001 Instructivo Carga de Nuevo Proyecto en Sistema Netsuite.docx", "output": "INST-COM-001 Instructivo Carga de Nuevo Proyecto V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-001", "input": "Comercial/PROC-COM-001 Análisis de Mercado V.2 2024.docx", "output": "PROC-COM-001 Análisis de Mercado V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-002", "input": "Comercial/PROC-COM-002 Gestión de Acuerdos Comerciales V.2.docx", "output": "PROC-COM-002 Gestión de Acuerdos Comerciales V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-003", "input": "Comercial/PROC-COM-003 Comunicación interna y con Cliente durante ejecución del Acuerdo Comercial V.2.docx", "output": "PROC-COM-003 Comunicación interna y con Cliente V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-004", "input": "Comercial/PROC-COM-004 Coordinación de operaciones portuarias de Carga y Descarga V.2.docx", "output": "PROC-COM-004 Coordinación de operaciones portuarias V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-005", "input": "Comercial/PROC-COM-005 Análisis para la Facturacion Comercial V.2.docx", "output": "PROC-COM-005 Análisis para la Facturación Comercial V.3.docx", "area": "Comercial"},
    {"code": "PROC-COM-006", "input": "Comercial/PROC-COM-006 Reportes Comerciales V.2.docx", "output": "PROC-COM-006 Reportes Comerciales V.3.docx", "area": "Comercial"},
    
    # OPERACIONES
    {"code": "PROC-OPS-002", "input": "Operaciones/PROC-OPS-002 Programación de Viaje V.2.docx", "output": "PROC-OPS-002 Programación de Viaje V.3.docx", "area": "Operaciones"},
    {"code": "PROC-OPS-003", "input": "Operaciones/PROC-OPS-003 Gestión de Operaciones V.2.docx", "output": "PROC-OPS-003 Gestión de Operaciones V.3.docx", "area": "Operaciones"},
    {"code": "PROC-OPS-004", "input": "Operaciones/PROC-OPS-004 Gestión de Finalización de Viaje V.2.docx", "output": "PROC-OPS-004 Gestión de Finalización de Viaje V.3.docx", "area": "Operaciones"},
    {"code": "PROC-OPS-005", "input": "Operaciones/PROC-OPS-005 Pronostico y monitoreo de precipitaciones y niveles de agua_V.1.docx", "output": "PROC-OPS-005 Pronóstico y monitoreo de precipitaciones V.3.docx", "area": "Operaciones"},
    
    # RRHH - Updated with all existing files
    {"code": "PROC-RHU-003", "input": "RRHH/PROC-RHU-003-Reclutamiento seleccion y contratacion del Personal.pdf", "output": "PROC-RHU-003 Reclutamiento selección y contratación del Personal V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-004", "input": "RRHH/PROC-RHU-004-Administrar el Bienestar de los Trabajadores.docx", "output": "PROC-RHU-004 Administrar el Bienestar de los Trabajadores V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-006", "input": "RRHH/PROC-RHU-006 Administración del Personal.docx", "output": "PROC-RHU-006 Administración del Personal V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-007", "input": "RRHH/PROC-RHU-007 Administración de Seguro Médico.docx", "output": "PROC-RHU-007 Administración de Seguro Médico V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-008", "input": "RRHH/PROC-RHU-008 Suspensión de Contratos de Trabajo.docx", "output": "PROC-RHU-008 Suspensión de Contratos de Trabajo V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-009", "input": "Crewing/PROC RHU-009 Gestion de Convocatoria a Tripulación v.2-2.docx", "output": "PROC-RHU-009 Gestión de Convocatoria a Tripulación V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-011", "input": "Crewing/PROC RHU-011 Información para liquidación Sueldos de Tripulación v.2-2.docx", "output": "PROC-RHU-011 Información para liquidación Sueldos V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-012", "input": "RRHH/PROC-RHU-012-Gestión de insumos y mantenimiento de oficinas.docx", "output": "PROC-RHU-012 Gestión de insumos y mantenimiento V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-014", "input": "RRHH/PROC-RHU-014-Gestión Documental en oficina.docx", "output": "PROC-RHU-014 Gestión Documental en oficina V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-015", "input": "RRHH/PROC-RHU-015-Gestión de viajes.docx", "output": "PROC-RHU-015 Gestión de viajes V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-016", "input": "RRHH/PROC-RHU-016 Evaluación de Desempeño.docx", "output": "PROC-RHU-016 Evaluación de Desempeño V.3.docx", "area": "RRHH"},
    {"code": "PROC-RHU-017", "input": "RRHH/PROC-RHU-017-Gestión de Uso de Vehículos.docx", "output": "PROC-RHU-017 Gestión de Uso de Vehículos V.3.docx", "area": "RRHH"},
    
    # SCH (Supply Chain) - Updated with actual versions
    {"code": "PROC-SCH-001", "input": "SCH/PROC-SCH-001-Gestión de Compras V.3.docx", "output": "PROC-SCH-001 Gestión de Compras V.3.docx", "area": "SCH"},
    {"code": "PROC-SCH-002", "input": "SCH/PROC-SCH-002-Selección Evaluación y Control de Proveedores V.3.docx", "output": "PROC-SCH-002 Selección Evaluación y Control de Proveedores V.3.docx", "area": "SCH"},
    {"code": "PROC-SCH-003", "input": "SCH/PROC-SCH-003 Recepcion de bienes y verificación de servicios V.2.docx", "output": "PROC-SCH-003 Recepción de bienes y verificación de servicios V.3.docx", "area": "SCH"},
    
    # IT
    {"code": "PROC-IT-001", "input": "IT/20251029 PROC-IT-001 Mantenimiento de Infraestructura IT-2.docx", "output": "PROC-IT-001 Mantenimiento de Infraestructura IT V.3.docx", "area": "IT"},
    
    # SEGURIDAD
    {"code": "PROC-BUQ-003", "input": "Seguridad/PROC-BUQ-003 Plan de Emergencia REV 3.docx", "output": "PROC-BUQ-003 Plan de Emergencia V.3.docx", "area": "Seguridad"},
    {"code": "PROC-GES-XXX", "input": "Seguridad/PROC-GES-XXX Gestión de Certificados.docx", "output": "PROC-GES-001 Gestión de Certificados V.3.docx", "area": "Seguridad"},
    
    # TECNICA
    {"code": "PROC-TEC-001", "input": "Tecnica/PROC-TEC-001 Mantenimiento de flota.docx", "output": "PROC-TEC-001 Mantenimiento de flota V.3.docx", "area": "Tecnica"},
]

def process_all_documents():
    """Process all documents in the inventory using SAFE OPTIMIZER"""
    
    base_dir = "/Users/lucas/Library/CloudStorage/GoogleDrive-altubelucas@gmail.com/Mi unidad/08 - ALTUBE IA/Interbarge/Sistema de Gestion por Procesos"
    
    # Create area folders in "Propuesta de Mejora Segura"
    areas = set([doc["area"] for doc in DOCUMENTS_INVENTORY])
    for area in areas:
        area_path = os.path.join(base_dir, "Propuesta de Mejora Segura", area)
        os.makedirs(area_path, exist_ok=True)
    
    print("🛡️  STARTING SAFE OPTIMIZATION OF ALL DOCUMENTS")
    print("⚠️  CONSERVATIVE MODE - Format preservation priority")
    print(f"📊 Total documents to process: {len(DOCUMENTS_INVENTORY)}")
    print("="*70)
    
    processed = 0
    errors = 0
    
    for doc_info in DOCUMENTS_INVENTORY:
        input_path = os.path.join(base_dir, doc_info["input"])
        output_path = os.path.join(base_dir, "Propuesta de Mejora Segura", doc_info["area"], doc_info["output"])
        
        print(f"\n[{processed + 1}/{len(DOCUMENTS_INVENTORY)}] Processing {doc_info['code']}...")
        
        # Check if input file exists and is a Word document
        if not os.path.exists(input_path):
            print(f"❌ ERROR: Input file not found: {input_path}")
            errors += 1
            continue
            
        # Skip PDF files (safe optimizer is for Word documents only)
        if input_path.lower().endswith('.pdf'):
            print(f"⚠️  SKIPPED: PDF file (not supported by safe optimizer)")
            continue
        
        # Run SAFE optimizer (not the dangerous universal one)
        try:
            result = subprocess.run([
                "python3", "safe_optimizer.py",
                input_path, output_path, doc_info["code"]
            ], capture_output=True, text=True, cwd=base_dir)
            
            if result.returncode == 0:
                processed += 1
                print("✅ SUCCESS")
            else:
                print(f"❌ ERROR: {result.stderr}")
                errors += 1
                
        except Exception as e:
            print(f"❌ EXCEPTION: {e}")
            errors += 1
    
    print("\n" + "="*70)
    print("🛡️  SAFE BATCH PROCESSING COMPLETED")
    print(f"✅ Successfully processed: {processed} documents")
    print(f"❌ Errors encountered: {errors} documents")
    print(f"📊 Success rate: {(processed/(processed+errors)*100):.1f}%")
    print("🔐 ALL format preservation protocols applied")
    
    return processed, errors

if __name__ == "__main__":
    process_all_documents()