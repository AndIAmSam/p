#!/usr/bin/env python3
import os
import sys
import argparse
import logging
import uuid
import datetime
import pandas as pd
import numpy as np
import unicodedata
import re
from sqlalchemy import create_engine, text

# Configurar el registro de logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# Función centralizada para normalizar números de póliza y evitar orfandad
def normalize_policy_number(val):
    if pd.isna(val) or str(val).lower() == 'nan':
        return None
    # Convertir a texto y limpiar espacios
    clean_val = str(val).strip()
    # Eliminar el '.0' al final si pandas lo leyó como float
    clean_val = re.sub(r'\.0$', '', clean_val)
    # Si es puramente numérico, quitar ceros a la izquierda
    if clean_val.isdigit():
        return str(int(clean_val))
    return clean_val

# Diccionario de homologación de Productos (Excel -> Base de datos)
PRODUCT_MAPPING = {
    # Gastos Médicos Mayores (GMM)
    'GMM EJECUTIVO': 'medicalife_familiar_ejecutivo',
    'GMM MAS': 'medicalife_familiar_mas',
    'GMM MAS ': 'medicalife_familiar_mas',
    'GMM  MAS': 'medicalife_familiar_mas',
    'GMM BASICO': 'medicalife_familiar_basico',
    'GMM BASICO ': 'medicalife_familiar_basico',
    'GMM  BASICO': 'medicalife_familiar_basico',
    'GMM CONVERSION': 'medicalife_familiar_conversion',
    
    # Vida
    'PRIMORDIAL': 'primordial',
    'PRIMORDIAL ': 'primordial',
    'METALIFE MUJER': 'metalife_mujer',
    'METALIFE MUJER ': 'metalife_mujer',
    'METALIFE TU FUTURO': 'metalife_tu_futuro',
    'METALIFE EDUCACION': 'metalife_educacion',
    'METALIFE EDUCACION ': 'metalife_educacion',
    'METALIFE RETIRO': 'metalife_retiro',
    'METALIFE RETIRO ': 'metalife_retiro',
    'PERFECTLIFE': 'perfeclife',
    'TEMPOLIFE': 'tempolife',
    'TEMPOLIFE 1': 'tempolife',
    'TEMPOLIFE 5': 'tempolife',
    'TEMPOLIFE 10': 'tempolife',
    'TEMPOLIFE 10 AÑOS': 'tempolife',
    'TEMPOLIFE 15 AÑOS': 'tempolife',
    'TEMPOLIFE 20 AÑOS': 'tempolife',
    'TEMPOLIFE GRANDES SUMAS 20': 'tempolife_grandes_sumas',
    'TOTALIFE': 'totalife',
    'TOTALIFE ': 'totalife',
    'TOTALIFE 5': 'totalife',
    'TOTALIFE 10': 'totalife',
    'TOTALIFE 10 ': 'totalife',
    'TOTALIFE 10 VIDAS CONJUNTAS': 'totalife',
    'TOTALIFE 15': 'totalife',
    'TOTALIFE 20': 'totalife',
    'HORIZONTE': 'horizonte',
    'HORIZONTE PL 10': 'horizonte',
    'HORIZONTE 10 PPR': 'horizonte',
    'HORIZONTE PPR': 'horizonte',
    'HORIZONTE PAGOS LIMITADOS': 'horizonte',
    'FLEXILIFE INVERSION': 'flexilife_inversion',
    'FLEXILIFE INVERSION  ': 'flexilife_inversion',
    'FLEXILIFE INVERSION 10': 'flexilife_inversion',
    'FLEXI LIFE INVERSION': 'flexilife_inversion',
    'FLEXILIFE': 'flexilife_suenos',
    'FLEXILIFE ': 'flexilife_suenos',
    'FLEXILIFE SUEÑOS 20': 'flexilife_suenos',
    'EDUCALIFE': 'educalife',
    
    # Mapeos de corrección y abreviaturas
    'GMM': 'medicalife_familiar_basico',
    'GMM ': 'medicalife_familiar_basico',
    'GMMI. MEDICALIFE': 'medicalife_familiar_basico',
    'METALIFE FUTURO': 'metalife_tu_futuro',
    'PERFECLIFE': 'perfeclife',
    'TOTALIFE 10 AÑOS': 'totalife',
    'FLEXIBLE': 'flexilife_suenos',
    'TEMPOLIFE GRANDES SUMAS 10': 'tempolife_grandes_sumas',
    'TEMPOLIFE 1 AÑO': 'tempolife',
    'GMM  EJECUTIVO': 'medicalife_familiar_ejecutivo',
    'MEDICALIFE EJECUTIVO': 'medicalife_familiar_ejecutivo',
    'METALFE RETIRO': 'metalife_retiro',
    'METALIRE RETIRO': 'metalife_retiro',
    'HORIZONTE  PPR': 'horizonte',
    'METELIFE RETIRO ': 'metalife_retiro',
    'METELIFE RETIRO': 'metalife_retiro',
    'METALIFE': 'metalife_retiro',
    'GMM FAMILIAR': 'medicalife_familiar_basico',
    'METLIFE RETIRO': 'metalife_retiro',
    'GMM MEDICALIFE': 'medicalife_familiar_basico',
    'PERFECTLIFE EA': 'perfeclife',
}

# Homologación de Asesores (Errores tipográficos -> Nombre oficial en BD)
ADVISOR_HOMOLOGATION = {
    'VICTOR DANIEL CUEVAS CLEMENTE': 'VICTOR DANIEL CUEVAS CLEMENTE',
    'VICTOR DANIEL CUEVAS CLEMENTE ': 'VICTOR DANIEL CUEVAS CLEMENTE',
    'SILVIA SAMANO MACIAS': 'SILVIA SAMANO MACIAS',
    'ROSARIO MARTINEZ HERNANDEZ': 'ROSARIO MARTINEZ HERNANDEZ',
    'ROSA ICELA VELAZQUEZ CORTEZ': 'ROSA ICELA VELAZQUEZ CORTEZ',
    # Agregar más nombres según sea necesario durante el Dry-Run
}

# --- FUNCIONES DE LIMPIEZA Y SOPORTE ---

def clean_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def clean_currency(val):
    if pd.isna(val):
        return 0.00
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).replace('$', '').replace(',', '').strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.00

def clean_date(val):
    if pd.isna(val) or val == 0 or str(val).strip() == '0':
        return None
    try:
        dt = pd.to_datetime(val)
        if dt.year < 1980: # descartar fechas erróneas
            return None
        return dt.date()
    except Exception:
        return None

def clean_rfc(val):
    if pd.isna(val) or not str(val).strip():
        return "XAXX010101000"
    clean = str(val).upper().replace(" ", "").replace("-", "").strip()
    return clean[:13]

def clean_phone(val):
    if pd.isna(val) or str(val).strip() == '0' or val == 0 or not str(val).strip():
        return "0000000000"
    clean = str(val).replace(" ", "").replace("-", "").replace("(", "").replace(")", "").strip()
    return clean[:20]

def parse_month_promotoria(val):
    """Convierte cadenas como 'jul,2024' en la fecha '2024-07-01'"""
    val = clean_str(val)
    if not val:
        return None
    months_es = {
        'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12
    }
    try:
        parts = val.split(',')
        if len(parts) == 2:
            m_str = parts[0].strip().lower()
            y_str = parts[1].strip()
            month = months_es.get(m_str[:3], 1)
            year = int(y_str)
            return datetime.date(year, month, 1)
    except Exception:
        pass
    return None

def split_mexican_name(full_name):
    """Divide un nombre en Nombre1, Nombre2, Apellido Paterno y Materno"""
    if not full_name:
        return "", None, "", ""
    words = full_name.split()
    first_name = ""
    second_name = None
    paternal = ""
    maternal = ""
    
    if len(words) == 1:
        first_name = words[0]
    elif len(words) == 2:
        first_name = words[0]
        paternal = words[1]
    elif len(words) == 3:
        first_name = words[0]
        paternal = words[1]
        maternal = words[2]
    else:
        first_name = words[0]
        second_name = words[1]
        paternal = words[2]
        maternal = words[3]
        
    return first_name, second_name, paternal, maternal

def map_payment_methods(conducto):
    """Mapea el conducto de pago a payment_method y payment_channel de SQL.
    Handles all known values from both Concentrado and Cobranza sheets."""
    c = clean_str(conducto).upper()
    if not c:
        return 'spei', 'direct'
    # Compound values first (most specific match wins)
    if 'CTA.CLABE' in c or 'CTA. CLABE' in c:
        return 'clabe', 'automatic'
    if 'NEXTPAY/MSI' in c or 'NEXT_PAY_MSI' in c:
        return 'next_pay_msi', 'automatic'
    if 'PORTAL MSI' in c or 'PORTAL_MSI' in c:
        return 'portal_msi', 'automatic'
    if 'PORTAL DEBITO' in c:
        return 'debit_card', 'automatic'
    # Single-word matches (order matters: specific before generic)
    if 'CLABE' in c:
        return 'clabe', 'automatic'
    if 'TDD' in c:
        return 'debit_card', 'automatic'
    if c in ('DEBITO', 'DEBITO '):
        return 'debit_card', 'automatic'
    if c in ('CREDITO', 'CREDITO ', 'CREITO', 'CREDTIO'):
        return 'credit_card', 'automatic'
    if 'MSI' in c:
        return 'portal_msi', 'automatic'
    if 'NEXTPAY' in c:
        return 'next_pay', 'direct'
    if 'TRANSFERENCIA' in c:
        return 'spei', 'direct'
    if 'VENTANILLA' in c:
        return 'bank_teller', 'direct'
    if 'EFECTIVO' in c:
        return 'bank_teller', 'direct'
    if c in ('PORTAL', 'PORTAL '):
        return 'portal', 'direct'
    # Cobranza genérico
    if 'AUTOMATICO' in c:
        return 'clabe', 'automatic'
    if 'DIRECTO' in c or 'AGENTE' in c:
        return 'spei', 'direct'
    # Default seguro
    logging.warning(f"Conducto de pago no reconocido: '{conducto}'. Se asignará por defecto a spei/direct.")
    return 'spei', 'direct'

def extract_validity_dates(group, default_start=None):
    if default_start is None:
        default_start = datetime.date.today()
        
    v_start = None
    if 'VIGENCIA DESDE' in group.columns:
        v_start_s = group['VIGENCIA DESDE'].apply(clean_date).dropna()
        v_start = v_start_s.min() if not v_start_s.empty else None
    
    v_end = None
    if 'VIGENCIA HASTA' in group.columns:
        v_end_s = group['VIGENCIA HASTA'].apply(clean_date).dropna()
        v_end = v_end_s.max() if not v_end_s.empty else None
    
    if v_start is None:
        v_start = default_start
    if v_end is None:
        v_end = v_start + datetime.timedelta(days=365)
        
    return v_start, v_end

def compute_receipt_validity_fallback(anchor_date, receipt_index, total_receipts):
    """Calcula valid_from y valid_until para un recibo cuando el Excel no tiene VIGENCIA.
    Replica la lógica del backfill SQL del backend: divide el año en partes iguales."""
    from dateutil.relativedelta import relativedelta
    if not anchor_date or total_receipts <= 0:
        # Fallback absoluto: fecha dummy
        return datetime.date(2020, 1, 1), datetime.date(2021, 1, 1)
    step_months = 12 // total_receipts
    if step_months < 1:
        step_months = 1
    vf = anchor_date + relativedelta(months=receipt_index * step_months)
    vu = anchor_date + relativedelta(months=(receipt_index + 1) * step_months)
    return vf, vu

def normalize_name(name):
    """
    Normaliza nombres para comparación exacta: 
    convierte a mayúsculas, quita acentos y espacios extra.
    """
    if not name or pd.isna(name):
        return ""
    # Convertir a string y mayúsculas
    name_str = str(name).strip().upper()
    # Quitar acentos usando NFD (Descomposición Canónica)
    name_str = ''.join(c for c in unicodedata.normalize('NFD', name_str) if unicodedata.category(c) != 'Mn')
    # Remover espacios dobles
    name_str = re.sub(r'\s+', ' ', name_str)
    return name_str

# --- ALGORITMO DE DEPURACIÓN DE DESPAGOS ---

def process_and_collapse_receipts(df_cob):
    """
    Agrupa transacciones por PÓLIZA y #RECIBO para calcular la suma neta,
    eliminar duplicados y resolver los 'Despagos' (-1000 que anula +1000).
    """
    if 'POLIZA RENOVACION' in df_cob.columns and 'POLIZA' not in df_cob.columns:
        df_cob = df_cob.rename(columns={'POLIZA RENOVACION': 'POLIZA'})
    if 'FECHA DE PAGO AÑO 1' in df_cob.columns and 'FECHA DE PAGO' not in df_cob.columns:
        df_cob = df_cob.rename(columns={'FECHA DE PAGO AÑO 1': 'FECHA DE PAGO'})
        
    # BUGFIX 4: Prevenir KeyErrors si las hojas no tienen columnas base
    if '#RECIBO' not in df_cob.columns:
        df_cob['#RECIBO'] = 'R1'
    if 'PCA ACREDITABLE' not in df_cob.columns:
        df_cob['PCA ACREDITABLE'] = 0.00
        
    collapsed = []
    
    # BUGFIX 10: Normalizar el número de recibo antes de agrupar.
    # Algunas pólizas tienen 'R1D-R12D ANUAL' (pago/despago) y luego 'R1-R12 ANUAL' (nuevo pago).
    # Deben colapsarse en el mismo bloque para no duplicar registros.
    def normalize_receipt(r):
        r_str = str(r).upper().strip()
        import re
        match = re.search(r'R(\d+)', r_str)
        if match:
            return 'R' + match.group(1)
        return r_str
        
    df_cob['#RECIBO_NORMALIZADO'] = df_cob['#RECIBO'].apply(normalize_receipt)
    grouped = df_cob.groupby(['POLIZA', '#RECIBO_NORMALIZADO'])
    
    for (poliza, recibo), group in grouped:
        net_pca = group['PCA ACREDITABLE'].sum()
        
        # Ordenar cronológicamente por fecha de pago para tomar el registro final
        if 'FECHA DE PAGO' in group.columns:
            group_sorted = group.sort_values(by='FECHA DE PAGO', ascending=True)
            latest_row = group_sorted.iloc[-1]
        else:
            group_sorted = group
            latest_row = group.iloc[-1]
        
        status_raw = clean_str(latest_row.get('ESTATUS')).lower()
        
        if net_pca > 0:
            status = 'paid'
            real_payment_date = clean_date(latest_row.get('FECHA DE PAGO'))
            if not real_payment_date and 'FECHA DE PAGO' in group.columns:
                # Si es pagado pero no tiene fecha, usar alguna fecha del grupo
                dates = group_sorted['FECHA DE PAGO'].apply(clean_date).dropna()
                real_payment_date = dates.iloc[-1] if not dates.empty else None
        else:
            # Si el neto es cero o negativo, se considera cancelado o pendiente de pago
            # BUGFIX 3: Coincidencia más robusta para estados cancelados
            status = 'cancelled' if group['ESTATUS'].astype(str).str.lower().str.contains('cancel|despago', na=False, regex=True).any() else 'not_paid_yet'
            real_payment_date = None
            net_pca = 0.00
            
        client_payment_date = None
        if 'FECHA DE PAGO (Segun Cliente)' in group.columns:
            client_dates = group['FECHA DE PAGO (Segun Cliente)'].apply(clean_date).dropna()
            if not client_dates.empty:
                client_payment_date = client_dates.iloc[0]
            
        comments = None
        if 'Observaciones' in group.columns:
            comments = "; ".join(group['Observaciones'].dropna().astype(str).unique())
        
        # Mapeos de fechas contables
        # BUGFIX 9: Para Mes Promotoria, tomar el PRIMER registro positivo.
        # Para QesQ, tomar el ÚLTIMO registro positivo.
        pos_group = group[group['PCA ACREDITABLE'] > 0] if 'PCA ACREDITABLE' in group.columns else group
        if pos_group.empty:
            pos_group = group
            
        first_promotoria = pos_group['Mes Promotoria'].dropna().iloc[0] if 'Mes Promotoria' in pos_group.columns and not pos_group['Mes Promotoria'].dropna().empty else None
        # BUGFIX 13: Para QesQ, también usar el PRIMER registro positivo en vez del último. 
        # Esto mantiene el PCA en su mes original y evita que el total del mes del despago/repago 
        # se infle artificialmente al comparar mes a mes contra el Excel.
        first_qesq = pos_group['QesQ'].dropna().iloc[0] if 'QesQ' in pos_group.columns and not pos_group['QesQ'].dropna().empty else None
        
        promotoria_month = parse_month_promotoria(first_promotoria)
        advisor_credited_date = parse_month_promotoria(first_qesq)
        
        # Vigencia del recibo (valid_from / valid_until) — extraer del Excel de Cobranza
        valid_from = None
        valid_until = None
        if 'VIGENCIA DESDE' in group.columns:
            vf_dates = group['VIGENCIA DESDE'].apply(clean_date).dropna()
            if not vf_dates.empty:
                valid_from = vf_dates.iloc[0]
        if 'VIGENCIA HASTA' in group.columns:
            vu_dates = group['VIGENCIA HASTA'].apply(clean_date).dropna()
            if not vu_dates.empty:
                valid_until = vu_dates.iloc[0]
        
        collapsed.append({
            'POLIZA': poliza,
            '#RECIBO': recibo,
            'status': status,
            'pca_acreditable': net_pca,
            'client_payment_date': client_payment_date,
            'real_payment_date': real_payment_date,
            'advisor_credited_date': advisor_credited_date,
            'promotoria_month': promotoria_month,
            'valid_from': valid_from,
            'valid_until': valid_until,
            'comments': comments[:500] if comments else None
        })
        
    df_result = pd.DataFrame(collapsed)
    
    # Ordenar recibos numéricamente (R1, R2, ... R10, R11) en vez de alfabéticamente (R1, R10, R11, R12, R2...)
    if not df_result.empty and '#RECIBO' in df_result.columns:
        df_result['_recibo_num'] = df_result['#RECIBO'].str.extract(r'(\d+)').astype(float)
        df_result = df_result.sort_values(['POLIZA', '_recibo_num']).drop(columns=['_recibo_num']).reset_index(drop=True)
    
    return df_result

# --- PROGRAMA PRINCIPAL ---

def main():
    parser = argparse.ArgumentParser(description="Script de Migración de Excel a MySQL para Pólizas y Recibos")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--dry-run', action='store_true', help="Ejecuta la migración en modo de prueba (valida todo pero no hace cambios permanentes)")
    group.add_argument('--live', action='store_true', help="Ejecuta la migración real aplicando cambios a la base de datos")
    
    parser.add_argument('--clean', action='store_true', help="Limpia las tablas de migración ANTES de insertar (elimina datos previos de migración)")
    parser.add_argument('--cobranza', default='datos_cobranza.xlsx', help="Path del archivo de cobranza")
    parser.add_argument('--concentrado', default='datos_concentrado.xlsx', help="Path del archivo de operaciones Metlife")
    parser.add_argument('--db-url', default=os.getenv('DATABASE_URL'), help="URL de conexión de base de datos (por defecto toma la variable de entorno DATABASE_URL)")
    
    args = parser.parse_args()
    
    if not args.db_url:
        logging.error("Falta la URL de base de datos. Pásala con --db-url o establece la variable de entorno DATABASE_URL.")
        sys.exit(1)
    
    if args.dry_run:
        logging.info("=== INICIANDO EN MODO DRY-RUN (SIMULACIÓN DE VALIDACIÓN) ===")
    else:
        logging.info("=== INICIANDO EN MODO LIVE (CARGA REAL DE DATOS) ===")
    
    # Pre-migration cleanup (if --clean flag is set)
    if args.clean:
        if 'prod' in args.db_url.lower():
            logging.error("SEGURIDAD: Estás intentando usar --clean contra una base de datos que parece ser de producción (la URL contiene 'prod'). Abortando por seguridad.")
            sys.exit(1)
            
        logging.info("Limpiando tablas de migración previas...")
        cleanup_engine = create_engine(args.db_url)
        with cleanup_engine.begin() as conn:
            conn.execute(text("SET FOREIGN_KEY_CHECKS = 0"))
            for tbl in ['policy_receipts', 'policy_renewals', 'policies', 'policy_applications']:
                conn.execute(text(f"TRUNCATE TABLE {tbl}"))
                logging.info(f"  Tabla '{tbl}' limpiada.")
            conn.execute(text("SET FOREIGN_KEY_CHECKS = 1"))
        cleanup_engine.dispose()
        logging.info("Limpieza completada. Las tablas están vacías y listas para la migración.")
        
    # Idempotency safeguard (if running live without clean)
    if args.live and not args.clean:
        try:
            check_engine = create_engine(args.db_url)
            with check_engine.connect() as conn:
                count = conn.execute(text("SELECT COUNT(*) FROM policies")).fetchone()[0]
                if count > 0:
                    logging.error(f"PROTECCIÓN DE IDEMPOTENCIA: La base de datos ya contiene {count} pólizas. Correr el script en modo --live sin --clean duplicaría masivamente los datos. Abortando por seguridad.")
                    sys.exit(1)
            check_engine.dispose()
        except Exception as e:
            logging.warning(f"No se pudo verificar la protección de idempotencia: {e}")
    # 1. Cargar Archivos Excel
    logging.info("Cargando archivos Excel...")
    if not os.path.exists(args.cobranza) or not os.path.exists(args.concentrado):
        logging.error(f"Archivos no encontrados en la raíz. Cobranza: {args.cobranza}, Concentrado: {args.concentrado}")
        sys.exit(1)
        
    try:
        # Cargar Venta Nueva (skipping metadata)
        df_vn = pd.read_excel(args.cobranza, sheet_name='Venta Nueva', skiprows=2)
        df_vn.columns = df_vn.columns.str.strip()
        
        # Cargar Concentrado Metlife (skiprows=0)
        df_con = pd.read_excel(args.concentrado, sheet_name=0)
        df_con.columns = df_con.columns.astype(str).str.strip()
        
        # Cargar Renovaciones Vida (skiprows=2)
        df_ren_vida = pd.read_excel(args.cobranza, sheet_name='RENOVACIONES VIDA', skiprows=2)
        df_ren_vida.columns = df_ren_vida.columns.str.strip()
        
        # Cargar Renovaciones GMM (skiprows=4)
        df_ren_gmm = pd.read_excel(args.cobranza, sheet_name='RENOVACIONES GMM', skiprows=4)
        df_ren_gmm.columns = df_ren_gmm.columns.str.strip()
        
        logging.info("Archivos cargados con éxito en memoria.")
    except Exception as e:
        logging.error(f"Error cargando las hojas de Excel: {e}")
        sys.exit(1)
        
    # 2. Pre-procesamiento e Imputación
    logging.info("Ejecutando algoritmos de normalización e imputación...")
    
    # Limpiar columnas clave de cruce (Venta Nueva y Concentrado)
    df_vn = df_vn.dropna(subset=['POLIZA'])
    df_vn['POLIZA'] = df_vn['POLIZA'].apply(normalize_policy_number)
    
    df_con = df_con.dropna(subset=['POLIZA'])
    df_con['POLIZA'] = df_con['POLIZA'].apply(normalize_policy_number)
    # BUGFIX 1: Deduplicar Concentrado por POLIZA para evitar multiplicar filas de VN en el merge
    df_con = df_con.drop_duplicates(subset='POLIZA', keep='first')
    
    # Normalizar Renovaciones GMM
    logging.info("Autocompletando datos dispersos en GMM...")
    # BUGFIX: Eliminamos la forzada a numérico para no destruir folios alfanuméricos en GMM
    df_ren_gmm['POLIZA RENOVACION'] = df_ren_gmm['POLIZA RENOVACION'].apply(normalize_policy_number)
    df_ren_gmm = df_ren_gmm.dropna(subset=['POLIZA RENOVACION'])
    
    # Rellenar valores nulos agrupando por POLIZA RENOVACION (GMM cambia folio anualmente)
    cols_to_fill = ['POLIZA ANTERIOR', 'AÑO VIDA POLIZA', 'PRODUCTO', 'Contratante', 'RFC', 'Asesor', 'FORMA DE PAGO', 'Conducto de Pago']
    for col in cols_to_fill:
        if col in df_ren_gmm.columns:
            df_ren_gmm[col] = df_ren_gmm.groupby('POLIZA RENOVACION')[col].transform(lambda x: x.ffill().bfill())
            
    # Default null AÑO VIDA POLIZA to 2 for GMM
    # BUGFIX 7: No podemos hacer fillna(2.0) aquí porque rompe el encadenamiento si la póliza es año 3 o superior.
    df_ren_gmm['AÑO VIDA POLIZA'] = pd.to_numeric(df_ren_gmm['AÑO VIDA POLIZA'], errors='coerce')
    df_ren_gmm['POLIZA ANTERIOR'] = df_ren_gmm['POLIZA ANTERIOR'].apply(normalize_policy_number)
    
    # BUGFIX 15: Si POLIZA ANTERIOR está vacía, usar POLIZA RENOVACION como póliza padre para no perder $1.4M de PCA
    # Debido a la normalización, nulos reales vienen como None
    df_ren_gmm.loc[df_ren_gmm['POLIZA ANTERIOR'].isna(), 'POLIZA ANTERIOR'] = df_ren_gmm['POLIZA RENOVACION']
            
    # Normalizar renovaciones Vida
    df_ren_vida['POLIZA RENOVACION'] = df_ren_vida['POLIZA RENOVACION'].apply(normalize_policy_number)
    df_ren_vida['POLIZA ANTERIOR'] = df_ren_vida['POLIZA ANTERIOR'].apply(normalize_policy_number)
    df_ren_vida = df_ren_vida.dropna(subset=['POLIZA ANTERIOR', 'POLIZA RENOVACION'])
    
    df_ren_vida['AÑO VIDA POLIZA'] = pd.to_numeric(df_ren_vida['AÑO VIDA POLIZA'], errors='coerce').fillna(2.0)
    df_ren_vida['AÑO VIDA POLIZA'] = df_ren_vida['AÑO VIDA POLIZA'].astype(int)
    
    # 3. Conciliación y Cruce de Datos (Venta Nueva + Concentrado)
    logging.info("Realizando cruce de datos Venta Nueva con Concentrado Metlife...")
    df_merged = pd.merge(df_vn, df_con, on='POLIZA', how='left', suffixes=('', '_con'))
    
    # PRIORIDAD: Datos de Cobranza (Venta Nueva) ganan. Concentrado solo se usa como fallback
    # cuando Cobranza tiene nulos. Esto se debe a que el Concentrado tiene más errores de captura.
    #
    # Las 7 columnas con nombre idéntico en ambos archivos generan sufijo '_con' al hacer el merge.
    # Para cada una, Cobranza (sin sufijo) se prioriza y solo se rellena con Concentrado (_con) si es nulo.
    if 'PRODUCTO_con' in df_merged.columns:
        df_merged['PRODUCTO'] = df_merged['PRODUCTO'].fillna(df_merged['PRODUCTO_con'])
    if 'CONTRATANTE' in df_merged.columns:
        df_merged['Contratante'] = df_merged['Contratante'].fillna(df_merged['CONTRATANTE'])
    if 'RAMO_con' in df_merged.columns:
        df_merged['RAMO'] = df_merged['RAMO'].fillna(df_merged['RAMO_con'])
    if 'RFC_con' in df_merged.columns:
        df_merged['RFC'] = df_merged['RFC'].fillna(df_merged['RFC_con'])
    if 'FORMA DE PAGO_con' in df_merged.columns:
        df_merged['FORMA DE PAGO'] = df_merged['FORMA DE PAGO'].fillna(df_merged['FORMA DE PAGO_con'])
    # FECHA DE PAGO: Esta columna existe en AMBOS archivos, pero para recibos debe venir EXCLUSIVAMENTE de Cobranza.
    # El Concentrado tiene una fecha genérica a nivel póliza que NO corresponde a las fechas reales de pago por recibo.
    # NO hacer fillna con FECHA DE PAGO_con — solo convertir a datetime para evitar errores de tipo al ordenar.
    if 'FECHA DE PAGO' in df_merged.columns:
        df_merged['FECHA DE PAGO'] = pd.to_datetime(df_merged['FECHA DE PAGO'], errors='coerce')
    # Convertir también FECHA DE PAGO (Segun Cliente), QesQ, Mes Promotoria — todas exclusivas de Cobranza
    if 'FECHA DE PAGO (Segun Cliente)' in df_merged.columns:
        df_merged['FECHA DE PAGO (Segun Cliente)'] = pd.to_datetime(df_merged['FECHA DE PAGO (Segun Cliente)'], errors='coerce')
    # ESTATUS: VN tiene estatus a nivel de RECIBO ('PAGADO','CANCELADO','DESPAGO'), CON tiene estatus a nivel de PÓLIZA
    # ('PAGADO','PENDIENTE','EN EMISION','NO TOMADA','CANCELADA'). Son conceptos diferentes pero ambos se priorizan desde Cobranza.
    # El sufijo automático del merge ya lo separa: 'ESTATUS' = VN, 'ESTATUS_con' = CON.
    # Conducto de Pago: Cobranza lo tiene como 'Conducto de Pago', Concentrado como 'CONDUCTO DE PAGO' (diferente case, no colisionan)
    if 'CONDUCTO DE PAGO' in df_merged.columns and 'Conducto de Pago' in df_merged.columns:
        df_merged['Conducto de Pago'] = df_merged['Conducto de Pago'].fillna(df_merged['CONDUCTO DE PAGO'])
    # Observaciones: Cobranza tiene 'Observaciones' (case mixto), Concentrado tiene 'OBSERVACIONES' (mayúscula). No colisionan.
    # Ambas se mantienen intactas; la de Cobranza se usa para recibos, la de Concentrado no se usa actualmente.
    
    logging.info("Prioridad aplicada: Todos los campos compartidos priorizan Cobranza sobre Concentrado.")
    
    # 4. Fase de Validación / Dry-Run (Detección de Inconsistencias)
    errors = []
    
    # Validar productos
    all_products = set(df_merged['PRODUCTO'].dropna().unique()).union(
        set(df_ren_vida['PRODUCTO'].dropna().unique())
    ).union(
        set(df_ren_gmm['PRODUCTO'].dropna().unique())
    )
    
    for prod in all_products:
        p_clean = str(prod).strip().upper()
        if p_clean not in PRODUCT_MAPPING:
            errors.append(f"Producto inválido o sin mapear en base de datos: '{prod}'")
            
    # Validar ramos
    all_branches = set(df_merged['RAMO'].dropna().unique()).union(
        set(df_ren_vida['RAMO'].dropna().unique())
    ).union(
        set(df_ren_gmm['RAMO'].dropna().unique())
    )
    for b in all_branches:
        b_clean = str(b).strip().upper()
        if b_clean not in ['VIDA', 'GMM', 'VIDA ', ' VIDA', 'GMM ']:
            errors.append(f"Ramo inválido: '{b}'")
            
    if errors:
        logging.error("--- SE DETECTARON ERRORES DE INTEGRIDAD DE DATOS ---")
        for err in errors[:20]:
            logging.error(err)
        if len(errors) > 20:
            logging.error(f"... y {len(errors) - 20} errores más.")
        logging.error("Migración cancelada. Limpia los datos de los archivos Excel e intenta de nuevo.")
        sys.exit(1)
        
    logging.info("Pre-validación completada: Cero errores críticos de mapeo de productos/ramos.")
    
    # 5. Iniciar Carga en Base de Datos
    engine = create_engine(args.db_url)
    
    try:
        with engine.begin() as conn: # Inicia transacción SQL (Auto-commit al salir, rollback en excepciones)
            logging.info("Abriendo transacción SQL...")
            
            # --- FASE 1: PROCESAR USUARIOS (ASESORES) ---
            logging.info("Procesando asesores comerciales...")
            
            # Cargar TODOS los asesores existentes y normalizarlos para evitar duplicados.
            # BUGFIX: Solo buscar contra roles de asesores (2=TeamLeader, 3=Advisor, 4=Advisor_Trainee)
            # y que no estén eliminados lógicamente, para evitar sobrescribir SuperAdmins/HR.
            existing_users = conn.execute(
                text("SELECT id, name FROM users WHERE deleted_at IS NULL AND role_id IN (2, 3, 4)")
            ).fetchall()
            db_users_map = {}
            for u_id, u_name in existing_users:
                norm_name = normalize_name(u_name)
                db_users_map[norm_name] = u_id
                
            # Unir asesores únicos de todas las hojas
            all_advisors = pd.concat([
                df_merged['Asesor'].dropna(),
                df_ren_vida['Asesor'].dropna(),
                df_ren_gmm['Asesor'].dropna()
            ]).str.strip().unique()
            
            # Extraer mapa de Clave Asesor desde Cobranza (para enriquecer advisor_key)
            advisor_key_map = {}  # Nombre Asesor -> Clave numérica
            for _, cob_df in [('VN', df_merged), ('Vida', df_ren_vida), ('GMM', df_ren_gmm)]:
                if 'Clave Asesor' in cob_df.columns:
                    for _, row in cob_df[['Asesor', 'Clave Asesor']].drop_duplicates().dropna(subset=['Asesor']).iterrows():
                        key_val = pd.to_numeric(row.get('Clave Asesor'), errors='coerce')
                        if pd.notna(key_val) and int(key_val) > 0:
                            advisor_key_map[str(row['Asesor']).strip()] = int(key_val)
            
            advisor_db_map = {} # Nombre Original -> user_id
            
            matched_advisors = []
            created_advisors = []
            
            for adv_name in all_advisors:
                # Homologar nombre
                clean_name = ADVISOR_HOMOLOGATION.get(adv_name, adv_name)
                norm_excel_name = normalize_name(clean_name)
                
                # Buscar si ya existe el asesor en la BD usando el mapa normalizado
                if norm_excel_name in db_users_map:
                    u_id = db_users_map[norm_excel_name]
                    matched_advisors.append(f"Excel: '{adv_name}' -> BD: Usuario ID {u_id}")
                else:
                    # Crear usuario temporal
                    # Añadir sufijo para garantizar unicidad y evitar error de clave única en BD
                    suffix = uuid.uuid4().hex[:4]
                    email_temp = f"{norm_excel_name.replace(' ', '.')}.{suffix}+migrado@vigvita.local"
                    conn.execute(
                        text("INSERT INTO users (name, email, password_hash, role_id) VALUES (:name, :email, 'temp_hash', 3)"),
                        {"name": clean_name, "email": email_temp}
                    )
                    # Recuperar ID
                    u_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                    # Insertar detalles de asesor con Clave del archivo de Cobranza (si disponible)
                    adv_key_val = advisor_key_map.get(adv_name, 0)
                    conn.execute(
                        text("INSERT INTO advisor_details (user_id, advisor_key, status) VALUES (:user_id, :key, 'pre_connected')"),
                        {"user_id": u_id, "key": adv_key_val}
                    )
                    created_advisors.append(f"'{adv_name}' (Email: {email_temp})")
                    # Añadir al mapa local para futuros matches
                    db_users_map[norm_excel_name] = u_id
                    
                # Si el asesor ya existía pero no tenía clave, actualizar con la de Cobranza
                adv_key_from_cob = advisor_key_map.get(adv_name)
                if adv_key_from_cob and adv_key_from_cob > 0:
                    res_upd = conn.execute(
                        text("UPDATE advisor_details SET advisor_key = :key WHERE user_id = :uid AND (advisor_key IS NULL OR advisor_key = 0)"),
                        {"key": adv_key_from_cob, "uid": u_id}
                    )
                    if res_upd.rowcount > 0:
                        logging.warning(f"Se actualizará el advisor_key del usuario existente ID {u_id} a {adv_key_from_cob}")
                    
                advisor_db_map[adv_name] = u_id
                
            logging.info("--- REPORTE DE MATCHING DE ASESORES ---")
            logging.info(f"Asesores cruzados exitosamente con usuarios existentes ({len(matched_advisors)}):")
            for m in matched_advisors:
                logging.info(f"  [MATCH] {m}")
            logging.info(f"Nuevos asesores 'fantasma' que se van a crear ({len(created_advisors)}):")
            for c in created_advisors:
                logging.info(f"  [NEW] {c}")
            logging.info("---------------------------------------")

            # BUGFIX 5: Obtener ID de fallback seguro creando un usuario huérfano explícito para evitar usar SuperAdmin
            res_fb = conn.execute(text("SELECT id FROM users WHERE email = 'huerfano.migracion@vigvita.local'")).fetchone()
            if not res_fb:
                conn.execute(text("INSERT INTO users (name, email, password_hash, role_id) VALUES ('Asesor Huérfano Migración', 'huerfano.migracion@vigvita.local', 'hash', 3)"))
                safe_fallback_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                # Insertar advisor_details para el huérfano
                conn.execute(
                    text("INSERT INTO advisor_details (user_id, advisor_key, status) VALUES (:user_id, 0, 'pre_connected')"),
                    {"user_id": safe_fallback_id}
                )
            else:
                safe_fallback_id = res_fb[0]
                
            logging.info(f"Asesores vinculados en base de datos. Total: {len(advisor_db_map)}")
            
            # --- FASE 2: PROCESAR PÓLIZAS DE VENTA NUEVA ---
            logging.info("Procesando pólizas de Venta Nueva...")
            policy_vn_db_map = {} # NumeroPoliza -> policy_id
            policy_validity_map = {} # NumeroPoliza -> validity_start (para fallback de vigencia de recibos)
            
            # Agrupar Venta Nueva por póliza para extraer el detalle del cliente
            grouped_vn = df_merged.groupby('POLIZA')
            for poliza_num, group in grouped_vn:
                # BUGFIX 12: La normalización alfanumérica y eliminación de ceros ahora se hace a nivel global
                # en el pre-procesamiento mediante normalize_policy_number() para evitar orfandad.
                clean_pol_num = poliza_num
                
                # Buscar fila maestra
                latest_master = group.iloc[-1]
                
                # Resolver campos: PRIORIDAD Cobranza > Concentrado (fallback)
                # Email y Teléfono solo existen en el Concentrado
                email = clean_str(latest_master.get('CORREO ELECTRONICO', latest_master.get('CORREO ELECTRÓNICO', '')))
                if not email:
                    email = "temporal@correo.com"
                
                phone = clean_phone(latest_master.get('No. DE CONTACTO', ''))
                # RFC: Ya priorizado en el merge (Cobranza gana)
                rfc = clean_rfc(latest_master.get('RFC', ''))
                    
                # Primas: Priorizar PNA de Cobranza, caer a Concentrado solo si no hay dato
                pna_cobranza = clean_currency(latest_master.get('PNA', 0.00))
                basic_con = clean_currency(latest_master.get('P. BASICA', latest_master.get('P. BÁSICA', 0.00)))
                add_prem = clean_currency(latest_master.get('P. ADICIONAL', 0.00))
                net_con = clean_currency(latest_master.get('PNA PESOS', 0.00))
                
                # PNA de Cobranza es la Prima Neta Anual, equivale a basic_premium + additional_premium
                basic_prem = pna_cobranza if pna_cobranza > 0 else (basic_con if basic_con > 0 else 0.00)
                net_prem = pna_cobranza if pna_cobranza > 0 else (net_con if net_con > 0 else basic_prem)
                
                # PCA: Usar PCA DEVENGADO de Cobranza (suma del grupo) como primera opción, caer a PCA del Concentrado
                pca_devengado_cobranza = clean_currency(latest_master.get('PCA DEVENGADO', 0.00))
                pca_concentrado = clean_currency(latest_master.get('PCA', latest_master.get('PCA ANUAL', 0.00)))
                pca_val = pca_devengado_cobranza if pca_devengado_cobranza > 0 else (pca_concentrado if pca_concentrado > 0 else net_prem)
                
                # Vigencia de la póliza completa
                validity_start, validity_end = extract_validity_dates(group, datetime.date.today())
                    
                # Split contratante
                contratante_raw = clean_str(latest_master['Contratante'])
                c_first, c_second, c_paternal, c_maternal = split_mexican_name(contratante_raw)
                
                # Estatus de Póliza — Se intenta derivar desde los recibos de Cobranza primero.
                # Si TODOS los recibos de VN para esta póliza son 'PAGADO', la póliza es 'issued'.
                # Si alguno es 'CANCELADO', podría ser 'suspended'. Pero VN no distingue entre
                # 'pending', 'in_emission' y 'not_taken' a nivel de póliza, así que para esos
                # estados especiales caemos al Concentrado (ESTATUS_con).
                status_map_policy = {
                    'PAGADO': 'issued',
                    'PENDIENTE': 'pending',
                    'EN EMISION': 'in_emission',
                    'EN EMISION ': 'in_emission',
                    'NO TOMADA': 'not_taken',
                    'NO TOMADO': 'not_taken',
                    'CANCELADA': 'suspended',
                }
                # Intentar primero con el ESTATUS_con del Concentrado (que es el estatus a nivel póliza)
                raw_status_con = clean_str(latest_master.get('ESTATUS_con', '')).upper().strip()
                status_vn = status_map_policy.get(raw_status_con, None)
                if status_vn is None:
                    # Si el Concentrado no tiene estatus, derivar de los recibos de Cobranza
                    receipt_statuses = group['ESTATUS'].dropna().str.upper().str.strip().unique()
                    if len(receipt_statuses) > 0:
                        if all('CANCELADO' in s or 'DESPAGO' in s for s in receipt_statuses):
                            status_vn = 'suspended'
                        else:
                            status_vn = 'issued'
                    else:
                        status_vn = 'issued'
                if status_vn not in ['pending', 'in_emission', 'suspended', 'not_taken', 'issued']:
                    status_vn = 'issued'
                    
                # Producto y Ramo
                prod_db = PRODUCT_MAPPING.get(clean_str(latest_master['PRODUCTO']).upper(), 'primordial')
                ramo_db = 'life' if 'VIDA' in clean_str(latest_master['RAMO']).upper() else 'major_medical'
                
                # Frecuencia e Impresion
                freq_db = 'annual'
                freq_raw = clean_str(latest_master['FORMA DE PAGO']).upper()
                if 'MENSUAL' in freq_raw:
                    freq_db = 'monthly'
                elif 'SEMESTRAL' in freq_raw:
                    freq_db = 'semiannual'
                elif 'TRIMESTRAL' in freq_raw:
                    freq_db = 'quarterly'
                elif 'BIMESTRAL' in freq_raw:
                    freq_db = 'bimonthly'
                    
                print_pol_val = 1 if 'SI' in clean_str(latest_master.get('IMPRESIÓN', 'NO')).upper() or 'IMPRESA' in clean_str(latest_master.get('IMPRESIÓN', 'NO')).upper() else 0
                
                # Resolver método de pago y canal: Priorizar Conducto de Pago de Cobranza
                conducto_cobranza = clean_str(latest_master.get('Conducto de Pago', ''))
                conducto_concentrado = clean_str(latest_master.get('CONDUCTO DE PAGO', ''))
                conducto_final = conducto_cobranza if conducto_cobranza else (conducto_concentrado if conducto_concentrado else 'Directo')
                method_db, channel_db = map_payment_methods(conducto_final)
                direct_debit_val = 1 if channel_db == 'automatic' else 0
                
                advisor_id = advisor_db_map.get(clean_str(latest_master['Asesor']), safe_fallback_id)
                
                # Generar UUIDs en Python para mayor seguridad de compatibilidad
                app_uuid = uuid.uuid4().bytes
                pol_uuid = uuid.uuid4().bytes
                folio_temp = f"VN-{clean_pol_num}"
                
                # 1. Crear policy_application
                conn.execute(
                    text("""
                        INSERT INTO policy_applications 
                        (public_id, created_by, application_folio, contractor_first_name, contractor_second_name, 
                         contractor_paternal_last_name, contractor_maternal_last_name, insurance_branch, product, 
                         payment_frequency, payment_method, basic_premium, additional_premium)
                        VALUES (:pid, :creator, :folio, :f_name, :s_name, :pat, :mat, :branch, :prod, :freq, :method, :basic, :add)
                    """),
                    {
                        "pid": app_uuid, "creator": advisor_id, "folio": folio_temp, "f_name": c_first, "s_name": c_second,
                        "pat": c_paternal, "mat": c_maternal, "branch": ramo_db, "prod": prod_db, "freq": freq_db,
                        "method": method_db, "basic": basic_prem, "add": add_prem
                    }
                )
                app_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                
                # 2. Crear policy
                conn.execute(
                    text("""
                        INSERT INTO policies 
                        (public_id, application_id, created_by, status, policy_number, application_folio, 
                         contractor_first_name, contractor_second_name, contractor_paternal_last_name, contractor_maternal_last_name, 
                         contractor_rfc, contractor_email, contractor_phone, print_policy, advisor_id, insurance_branch, 
                         product, inc_sum_assured, multiple_insured, payment_frequency, payment_method, payment_channel, 
                         direct_debit, pca, validity_start, validity_end, annual_basic_premium, annual_additional_premium, 
                         annual_net_premium, issue_date)
                        VALUES (:pid, :app_id, :creator, :status, :pol_num, :folio, :f_name, :s_name, :pat, :mat, 
                                :rfc, :email, :phone, :print_pol, :adv_id, :branch, :prod, :inc, :mult, :freq, :method, 
                                :channel, :dd, :pca, :v_start, :v_end, :basic, :add, :net, :issue)
                    """),
                    {
                        "pid": pol_uuid, "app_id": app_id, "creator": advisor_id, "status": status_vn, "pol_num": clean_pol_num,
                        "folio": folio_temp, "f_name": c_first, "s_name": c_second, "pat": c_paternal, "mat": c_maternal,
                        "rfc": rfc, "email": email, "phone": phone, "print_pol": print_pol_val, "adv_id": advisor_id,
                        "branch": ramo_db, "prod": prod_db, "inc": 0, "mult": 0, "freq": freq_db, "method": method_db,
                        "channel": channel_db, "dd": direct_debit_val, "pca": pca_val, "v_start": validity_start,
                        "v_end": validity_end, "basic": basic_prem, "add": add_prem, "net": net_prem, "issue": validity_start
                    }
                )
                pol_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                
                policy_vn_db_map[clean_pol_num] = pol_id
                policy_validity_map[clean_pol_num] = validity_start
                
            logging.info(f"Pólizas creadas. Total: {len(policy_vn_db_map)}")
            
            # --- FASE 3: CONSOLIDAR Y CARGAR RECIBOS DE VENTA NUEVA ---
            logging.info("Consolidando y limpiando despagos de cobranza Venta Nueva...")
            df_receipts_vn = process_and_collapse_receipts(df_merged)
            logging.info(f"Recibos a insertar tras depuración: {len(df_receipts_vn)}")
            
            # Agrupar recibos por póliza para calcular fallback de vigencia si es necesario
            for pol_num_vn, pol_group_vn in df_receipts_vn.groupby('POLIZA'):
                pol_key_vn = str(pol_num_vn).strip()
                if pol_key_vn.isdigit():
                    pol_key_vn = str(int(pol_key_vn))
                pol_id = policy_vn_db_map.get(pol_key_vn)
                if not pol_id:
                    continue
                anchor_vn = policy_validity_map.get(pol_key_vn)
                total_recs_vn = len(pol_group_vn)
                
                for seq_idx, (_, row) in enumerate(pol_group_vn.iterrows()):
                    vf = row['valid_from']
                    vu = row['valid_until']
                    if pd.isna(vf) or vf is None or pd.isna(vu) or vu is None:
                        vf, vu = compute_receipt_validity_fallback(anchor_vn, seq_idx, total_recs_vn)
                    
                    rec_uuid = uuid.uuid4().bytes
                    conn.execute(
                        text("""
                            INSERT INTO policy_receipts 
                            (public_id, policy_id, renewal_id, status, pca_acreditable, client_payment_date, real_payment_date, 
                             advisor_credited_date, promotoria_month, valid_from, valid_until, comments)
                            VALUES (:pid, :pol_id, NULL, :status, :pca, :c_date, :r_date, :adv_date, :prom_date, :vf, :vu, :comments)
                        """),
                        {
                            "pid": rec_uuid, "pol_id": pol_id, "status": row['status'], "pca": row['pca_acreditable'],
                            "c_date": row['client_payment_date'], "r_date": row['real_payment_date'],
                            "adv_date": row['advisor_credited_date'], "prom_date": row['promotoria_month'],
                            "vf": vf, "vu": vu,
                            "comments": row['comments']
                        }
                    )
                
            logging.info("Recibos de Venta Nueva cargados con éxito.")
            
            # --- FASE 4: PROCESAR RENOVACIONES (VIDA + GMM) ---
            logging.info("Procesando renovaciones de Vida y GMM...")
            
            # Unir dataframes de renovación
            df_ren_vida['RAMO_RENOVACION'] = 'life'
            df_ren_gmm['RAMO_RENOVACION'] = 'major_medical'
            
            # Asegurar tipos correctos para agrupaciones (ya procesados en Fase 2)
            pass
            
            # Procesar Vida
            logging.info("Cargando renovaciones de Vida...")
            grouped_ren_vida = df_ren_vida.groupby(['POLIZA ANTERIOR', 'AÑO VIDA POLIZA'])
            
            for (pol_anterior, ano_vida), group in grouped_ren_vida:
                latest_ren = group.iloc[-1]
                
                # 1. Recuperar ID de la póliza padre en base de datos
                res_pol = conn.execute(text("SELECT id, advisor_id, product FROM policies WHERE policy_number = :pol"), {"pol": pol_anterior}).fetchone()
                if not res_pol:
                    # Crear póliza base ficticia (Año 1) para mantener integridad
                    contratante_raw = clean_str(latest_ren.get('Contratante', 'CONTRATANTE HISTORICO'))
                    c_first, c_second, c_paternal, c_maternal = split_mexican_name(contratante_raw)
                    email = "temporal@correo.com"
                    phone = "0000000000"
                    rfc = clean_str(latest_ren.get('RFC', 'XAXX010101000'))
                    prod_db = PRODUCT_MAPPING.get(clean_str(latest_ren['PRODUCTO']).upper(), 'primordial')
                    
                    app_uuid = uuid.uuid4().bytes
                    pol_uuid = uuid.uuid4().bytes
                    advisor_id = advisor_db_map.get(clean_str(latest_ren['Asesor']), safe_fallback_id)
                    
                    conn.execute(
                        text("INSERT INTO policy_applications (public_id, created_by, application_folio, contractor_first_name, contractor_paternal_last_name, contractor_maternal_last_name, insurance_branch, product, payment_frequency, payment_method, basic_premium) VALUES (:pid, :creator, :folio, :first, :pat, :mat, 'life', :prod, 'annual', 'spei', 0.00)"),
                        {"pid": app_uuid, "creator": advisor_id, "folio": f"DV-{pol_anterior}", "first": c_first, "pat": c_paternal, "mat": c_maternal, "prod": prod_db}
                    )
                    app_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                    
                    conn.execute(
                        text("INSERT INTO policies (public_id, application_id, created_by, status, policy_number, application_folio, contractor_first_name, contractor_paternal_last_name, contractor_maternal_last_name, contractor_rfc, contractor_email, contractor_phone, print_policy, advisor_id, insurance_branch, product, inc_sum_assured, multiple_insured, payment_frequency, payment_method, payment_channel, direct_debit, pca, validity_start, validity_end, annual_basic_premium, annual_additional_premium, annual_net_premium, issue_date) VALUES (:pid, :app_id, :creator, 'issued', :pol_num, :folio, :first, :pat, :mat, :rfc, :email, :phone, 0, :adv_id, 'life', :prod, 0, 0, 'annual', 'spei', 'direct', 0, 0.00, '2020-01-01', '2021-01-01', 0.00, 0.00, 0.00, '2020-01-01')"),
                        {"pid": pol_uuid, "app_id": app_id, "creator": advisor_id, "pol_num": pol_anterior, "folio": f"DV-{pol_anterior}", "first": c_first, "pat": c_paternal, "mat": c_maternal, "rfc": rfc, "email": email, "phone": phone, "adv_id": advisor_id, "prod": prod_db}
                    )
                    padre_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                    advisor_id = advisor_id
                else:
                    padre_id, advisor_id, prod_db = res_pol
                    
                # 2. Insertar en policy_renewals
                validity_start, validity_end = extract_validity_dates(group, datetime.date.today())
                pca_val = clean_currency(latest_ren.get('PCA DEVENGADO', 0.00))
                
                # Mapeo de método de pago
                method_db, channel_db = map_payment_methods(latest_ren.get('Conducto de Pago', 'Directo'))
                
                # Frecuencia de pago desde la hoja de renovaciones
                freq_db_ren = 'annual'
                freq_raw_ren = clean_str(latest_ren.get('FORMA DE PAGO', '')).upper()
                if 'MENSUAL' in freq_raw_ren:
                    freq_db_ren = 'monthly'
                elif 'SEMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'semiannual'
                elif 'TRIMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'quarterly'
                elif 'BIMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'bimonthly'
                
                # Si es año 1, no creamos renovación en BD ya que es la póliza base original.
                # BUGFIX 11: Si la póliza ya fue procesada en Venta Nueva, NO volver a insertar
                # sus recibos. Solo procesar si es una póliza que NO está en VN (huérfana).
                if int(ano_vida) <= 1:
                    if str(pol_anterior) in policy_vn_db_map:
                        # Póliza ya procesada completamente en Venta Nueva — omitir para evitar duplicación
                        continue
                    ren_id = None
                else:
                    ren_uuid = uuid.uuid4().bytes
                    conn.execute(
                        text("""
                            INSERT INTO policy_renewals 
                            (public_id, policy_id, renewal_number, new_policy_number, advisor_id, inc_sum_assured, multiple_insured, 
                             payment_frequency, payment_method, payment_channel, direct_debit, pca, validity_start, validity_end, 
                             annual_basic_premium, annual_additional_premium, annual_net_premium)
                            VALUES (:pid, :pol_id, :ren_num, NULL, :adv_id, 0, 0, :freq, :method, :channel, :dd, :pca, :v_start, :v_end, :basic, 0.00, :basic)
                        """),
                        {
                            "pid": ren_uuid, "pol_id": padre_id, "ren_num": int(ano_vida) - 1, "adv_id": advisor_id,
                            "freq": freq_db_ren, "method": method_db, "channel": channel_db, "dd": 1 if channel_db == 'automatic' else 0,
                            "pca": pca_val, "v_start": validity_start, "v_end": validity_end, "basic": 0.00 # BUGFIX 2: Prima básica es 0.00, no PCA (que es comisión)
                        }
                    )
                    ren_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                
                # 3. Consolidar recibos de la renovación e insertar
                df_receipts_ren = process_and_collapse_receipts(group)
                total_recs_ren = len(df_receipts_ren)
                for seq_idx_ren, (_, r_row) in enumerate(df_receipts_ren.iterrows()):
                    vf_ren = r_row['valid_from']
                    vu_ren = r_row['valid_until']
                    if pd.isna(vf_ren) or vf_ren is None or pd.isna(vu_ren) or vu_ren is None:
                        vf_ren, vu_ren = compute_receipt_validity_fallback(validity_start, seq_idx_ren, total_recs_ren)
                    
                    rec_uuid = uuid.uuid4().bytes
                    conn.execute(
                        text("""
                            INSERT INTO policy_receipts 
                            (public_id, policy_id, renewal_id, status, pca_acreditable, client_payment_date, real_payment_date, 
                             advisor_credited_date, promotoria_month, valid_from, valid_until, comments)
                            VALUES (:pid, :pol_id, :ren_id, :status, :pca, :c_date, :r_date, :adv_date, :prom_date, :vf, :vu, :comments)
                        """),
                        {
                            "pid": rec_uuid, "pol_id": padre_id, "ren_id": ren_id, "status": r_row['status'], "pca": r_row['pca_acreditable'],
                            "c_date": r_row['client_payment_date'], "r_date": r_row['real_payment_date'],
                            "adv_date": r_row['advisor_credited_date'], "prom_date": r_row['promotoria_month'],
                            "vf": vf_ren, "vu": vu_ren,
                            "comments": r_row['comments']
                        }
                    )
                    
            logging.info("Renovaciones de Vida cargadas.")
            
            # Procesar GMM
            logging.info("Cargando renovaciones de GMM...")
            df_ren_gmm_clean = df_ren_gmm[(df_ren_gmm['POLIZA ANTERIOR'] != '0') & (df_ren_gmm['POLIZA ANTERIOR'].astype(str).str.lower() != 'nan')]
            # BUGFIX 6: Usar dropna=False y sort=False para procesar en orden de aparición (Año 2 aparece antes que Año 3)
            grouped_ren_gmm = df_ren_gmm_clean.groupby(['POLIZA ANTERIOR', 'AÑO VIDA POLIZA'], dropna=False, sort=False)
            
            # Conjunto para rastrear folios de renovación GMM ya insertados (evitar Violación de Clave Única)
            inserted_gmm_pol_nums = set()
            
            for (pol_anterior, ano_vida), group in grouped_ren_gmm:
                latest_ren = group.iloc[-1]
                
                # Obtener la póliza de renovación para este año (tomando la última no nula de este grupo)
                pol_ren_series = group['POLIZA RENOVACION'].dropna()
                pol_renovacion = pol_ren_series.iloc[-1] if not pol_ren_series.empty else None
                
                # 1. Buscar padre (primero en policies, luego en policy_renewals para encadenamiento de años > 2)
                res_pol = conn.execute(text("SELECT id, advisor_id FROM policies WHERE policy_number = :pol"), {"pol": pol_anterior}).fetchone()
                
                base_ren_num = 0
                # BUGFIX 6: Si no está en policies, quizás es una renovación previa (GMM encadena Año 3 -> Año 2)
                if not res_pol:
                    res_ren = conn.execute(text("SELECT policy_id, advisor_id, renewal_number FROM policy_renewals WHERE new_policy_number = :pol"), {"pol": pol_anterior}).fetchone()
                    if res_ren:
                        res_pol = (res_ren[0], res_ren[1])
                        base_ren_num = res_ren[2]
                        
                # Calcular dinámicamente el año de vida si venía vacío en el Excel
                if pd.isna(ano_vida):
                    ano_vida = base_ren_num + 2
                        
                if not res_pol:
                    # Crear póliza base ficticia (Año 1)
                    contratante_raw = clean_str(latest_ren.get('Contratante', 'CONTRATANTE HISTORICO'))
                    c_first, c_second, c_paternal, c_maternal = split_mexican_name(contratante_raw)
                    email = "temporal@correo.com"
                    phone = "0000000000"
                    rfc = clean_str(latest_ren.get('RFC', 'XAXX010101000'))
                    prod_db = PRODUCT_MAPPING.get(clean_str(latest_ren['PRODUCTO']).upper(), 'medicalife_familiar_basico')
                    
                    app_uuid = uuid.uuid4().bytes
                    pol_uuid = uuid.uuid4().bytes
                    advisor_id = advisor_db_map.get(clean_str(latest_ren['Asesor']), safe_fallback_id)
                    
                    conn.execute(
                        text("INSERT INTO policy_applications (public_id, created_by, application_folio, contractor_first_name, contractor_paternal_last_name, contractor_maternal_last_name, insurance_branch, product, payment_frequency, payment_method, basic_premium) VALUES (:pid, :creator, :folio, :first, :pat, :mat, 'major_medical', :prod, 'annual', 'spei', 0.00)"),
                        {"pid": app_uuid, "creator": advisor_id, "folio": f"DG-{pol_anterior}", "first": c_first, "pat": c_paternal, "mat": c_maternal, "prod": prod_db}
                    )
                    app_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                    
                    conn.execute(
                        text("INSERT INTO policies (public_id, application_id, created_by, status, policy_number, application_folio, contractor_first_name, contractor_paternal_last_name, contractor_maternal_last_name, contractor_rfc, contractor_email, contractor_phone, print_policy, advisor_id, insurance_branch, product, inc_sum_assured, multiple_insured, payment_frequency, payment_method, payment_channel, direct_debit, pca, validity_start, validity_end, annual_basic_premium, annual_additional_premium, annual_net_premium, issue_date) VALUES (:pid, :app_id, :creator, 'issued', :pol_num, :folio, :first, :pat, :mat, :rfc, :email, :phone, 0, :adv_id, 'major_medical', :prod, 0, 0, 'annual', 'spei', 'direct', 0, 0.00, '2020-01-01', '2021-01-01', 0.00, 0.00, 0.00, '2020-01-01')"),
                        {"pid": pol_uuid, "app_id": app_id, "creator": advisor_id, "pol_num": pol_anterior, "folio": f"DG-{pol_anterior}", "first": c_first, "pat": c_paternal, "mat": c_maternal, "rfc": rfc, "email": email, "phone": phone, "adv_id": advisor_id, "prod": prod_db}
                    )
                    padre_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                    advisor_id = advisor_id
                else:
                    padre_id, advisor_id = res_pol
                    
                # 2. Insertar en policy_renewals (GMM sí tiene new_policy_number)
                validity_start, validity_end = extract_validity_dates(group, datetime.date.today())
                pca_val = clean_currency(latest_ren.get('PCA DEVENGADO', 0.00))
                method_db, channel_db = map_payment_methods(latest_ren.get('Conducto de Pago', 'Directo'))
                
                # Frecuencia de pago desde la hoja de renovaciones
                freq_db_ren = 'annual'
                freq_raw_ren = clean_str(latest_ren.get('FORMA DE PAGO', '')).upper()
                if 'MENSUAL' in freq_raw_ren:
                    freq_db_ren = 'monthly'
                elif 'SEMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'semiannual'
                elif 'TRIMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'quarterly'
                elif 'BIMESTRAL' in freq_raw_ren:
                    freq_db_ren = 'bimonthly'
                
                # Si es año 1, no creamos renovación en BD ya que es la póliza base original.
                # BUGFIX 11: Si la póliza ya fue procesada en Venta Nueva, NO volver a insertar
                # sus recibos. Solo procesar si es una póliza que NO está en VN (huérfana).
                if int(ano_vida) <= 1:
                    if str(pol_anterior) in policy_vn_db_map:
                        # Póliza ya procesada completamente en Venta Nueva — omitir para evitar duplicación
                        continue
                    ren_id = None
                else:
                    ren_num_to_insert = int(ano_vida) - 1
                    
                    # BUGFIX 14: Verificar si esta renovación ya existe (puede haber sido insertada por Vida)
                    existing_ren = conn.execute(
                        text("SELECT id FROM policy_renewals WHERE policy_id = :pid AND renewal_number = :rn"),
                        {"pid": padre_id, "rn": ren_num_to_insert}
                    ).fetchone()
                    
                    if existing_ren:
                        # La renovación ya existe (insertada por Vida u otro grupo GMM anterior), reutilizar su ID
                        ren_id = existing_ren[0]
                    else:
                        new_pol_val = pol_renovacion
                        if new_pol_val is None:
                            new_pol_param = None
                        elif new_pol_val in inserted_gmm_pol_nums:
                            logging.warning(f"Número de renovación GMM duplicado detectado: {new_pol_val}. Se insertará con new_policy_number = NULL para conservar integridad y unicidad.")
                            new_pol_param = None
                        else:
                            new_pol_param = new_pol_val
                            inserted_gmm_pol_nums.add(new_pol_val)
                            
                        ren_uuid = uuid.uuid4().bytes
                        conn.execute(
                            text("""
                                INSERT INTO policy_renewals 
                                (public_id, policy_id, renewal_number, new_policy_number, advisor_id, inc_sum_assured, multiple_insured, 
                                 payment_frequency, payment_method, payment_channel, direct_debit, pca, validity_start, validity_end, 
                                 annual_basic_premium, annual_additional_premium, annual_net_premium)
                                VALUES (:pid, :pol_id, :ren_num, :new_pol, :adv_id, 0, 0, :freq, :method, :channel, :dd, :pca, :v_start, :v_end, :basic, 0.00, :basic)
                            """),
                            {
                                "pid": ren_uuid, "pol_id": padre_id, "ren_num": ren_num_to_insert, "new_pol": new_pol_param,
                                "adv_id": advisor_id, "freq": freq_db_ren, "method": method_db, "channel": channel_db, "dd": 1 if channel_db == 'automatic' else 0,
                                "pca": pca_val, "v_start": validity_start, "v_end": validity_end, "basic": 0.00 # BUGFIX 2: Prima básica es 0.00, no PCA (que es comisión)
                            }
                        )
                        ren_id = conn.execute(text("SELECT LAST_INSERT_ID()")).fetchone()[0]
                
                # 3. Consolidar recibos de la renovación e insertar
                df_receipts_ren = process_and_collapse_receipts(group)
                total_recs_ren_gmm = len(df_receipts_ren)
                for seq_idx_gmm, (_, r_row) in enumerate(df_receipts_ren.iterrows()):
                    vf_gmm = r_row['valid_from']
                    vu_gmm = r_row['valid_until']
                    if pd.isna(vf_gmm) or vf_gmm is None or pd.isna(vu_gmm) or vu_gmm is None:
                        vf_gmm, vu_gmm = compute_receipt_validity_fallback(validity_start, seq_idx_gmm, total_recs_ren_gmm)
                    
                    rec_uuid = uuid.uuid4().bytes
                    conn.execute(
                        text("""
                            INSERT INTO policy_receipts 
                            (public_id, policy_id, renewal_id, status, pca_acreditable, client_payment_date, real_payment_date, 
                             advisor_credited_date, promotoria_month, valid_from, valid_until, comments)
                            VALUES (:pid, :pol_id, :ren_id, :status, :pca, :c_date, :r_date, :adv_date, :prom_date, :vf, :vu, :comments)
                        """),
                        {
                            "pid": rec_uuid, "pol_id": padre_id, "ren_id": ren_id, "status": r_row['status'], "pca": r_row['pca_acreditable'],
                            "c_date": r_row['client_payment_date'], "r_date": r_row['real_payment_date'],
                            "adv_date": r_row['advisor_credited_date'], "prom_date": r_row['promotoria_month'],
                            "vf": vf_gmm, "vu": vu_gmm,
                            "comments": r_row['comments']
                        }
                    )
                    
            logging.info("Renovaciones de GMM cargadas con éxito.")
            
            # --- FINALIZAR TRANSACCIÓN ---
            if args.dry_run:
                # Simular error para forzar rollback en modo de prueba
                raise Exception("Mecanismo de seguridad DRY-RUN: Revirtiendo inserciones virtuales.")
            else:
                logging.info("=== MIGRACIÓN LIVE COMPLETADA CON ÉXITO. TRANSACCIÓN CONFIRMADA (COMMIT) ===")
                
    except Exception as e:
        if args.dry_run and "Mecanismo de seguridad DRY-RUN" in str(e):
            logging.info("=== SIMULACIÓN COMPLETADA SIN ERRORES EN BASE DE DATOS. REVERSIÓN LOGRADA CON ÉXITO (ROLLBACK) ===")
        else:
            logging.error(f"Ocurrió un error inesperado. Se ejecutó un ROLLBACK. Base de datos segura. Detalle: {e}")
            sys.exit(1)

if __name__ == '__main__':
    main()
