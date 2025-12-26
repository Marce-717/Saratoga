# -*- coding: utf-8 -*-
"""
Created on Sun Dec 21 13:16:48 2025

@author: marce
"""

import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

"""
LIMPIEZA DE DATOS: ventas2025.xlsx
==================================
Script para limpiar y preparar datos de ventas para análisis

Autor: Puelche
Fecha: Diciembre 2024
"""

ruta_nueva = 'C:/Users/marce/OneDrive/Documentos/OneDrive/Documentos/Saratoga'  # Especifica la ruta deseada aquí
os.chdir(ruta_nueva)  # Cambia el directorio de trabajo
nueva_ruta = os.getcwd()  # Obtiene la ruta actual para verificar el cambio
print(nueva_ruta)  # Imprime la nueva ruta
print("ESTADISTICAS DE VENTAS DE SARATOGA \n")
# Como experto en programación Phyton 3 quiero reemplazar el rango de datos de ventas[98:128] por 01-1
# # QUiero hacer un análisis exploratorio de variables, para ello quiero hacer un ciclo FOR de lista_categorica y lista_numerica


# ============================================================
# 1. CARGAR DATOS ORIGINALES
# ============================================================

print("ESTADISTICAS DE VENTAS DE SARATOGA \n")
# Como experto en programación Phyton 3 quiero reemplazar el rango de datos de ventas[98:128] por 01-1
# # QUiero hacer un análisis exploratorio de variables, para ello quiero hacer un ciclo FOR de lista_categorica y lista_numerica
df_original = pd.read_excel('ventas2025.xlsx', sheet_name= "hoja2025")

print(f"\n✓ Datos cargados: {df_original.shape[0]} filas x {df_original.shape[1]} columnas")
print("\n📋 COLUMNAS DEL ARCHIVO:")
print(df_original.columns.tolist())

# Crear copia para trabajar
df = df_original.copy()

# Renombrar columnas para trabajar más fácil (adaptado a tus nombres)
# Verificar qué columnas existen y renombrar según corresponda
columnas_map = {}
for col in df.columns:
    col_lower = col.lower().strip()
    if 'nro' in col_lower or 'número' in col_lower:
        columnas_map[col] = 'nro'
    elif 'ped' in col_lower and 'pedido' in col_lower:
        columnas_map[col] = 'pedido'
    elif 'tipo' in col_lower and 'doc' in col_lower:
        columnas_map[col] = 'Tipo_doc'
    elif col_lower == 'documento':
        columnas_map[col] = 'Documento'
    elif 'fecha' in col_lower and 'pedido' in col_lower:
        columnas_map[col] = 'fecha_pedido'
    elif 'fecha' in col_lower and 'documento' in col_lower:
        columnas_map[col] = 'fecha_documento'
    elif 'cliente' in col_lower:
        columnas_map[col] = 'Cliente'
    elif 'vendedor' in col_lower:
        columnas_map[col] = 'Vendedor'
    elif col_lower == 'cc':
        columnas_map[col] = 'CC'
    elif 'estado' in col_lower:
        columnas_map[col] = 'Estado'
    elif 'pag' in col_lower or 'pago' in col_lower:
        columnas_map[col] = 'F_Pag_Principal'
    elif 'neto' in col_lower:
        columnas_map[col] = 'Neto'
    elif 'iva' in col_lower:
        columnas_map[col] = 'IVA'
    elif 'bruto' in col_lower:
        columnas_map[col] = 'Bruto'
    elif 'servicio' in col_lower:
        columnas_map[col] = 'Servicios'
    elif 'total' in col_lower:
        columnas_map[col] = 'Total'

df = df.rename(columns=columnas_map)
print("\n✓ Columnas renombradas para análisis")
print(f"Columnas disponibles: {df.columns.tolist()}")

# ============================================================
# 2. DIAGNÓSTICO INICIAL
# ============================================================

print("\n" + "=" * 70)
print("DIAGNÓSTICO INICIAL DE VALORES FALTANTES")
print("=" * 70)

valores_na = df.isna().sum()
porcentaje_na = (df.isna().sum() / len(df)) * 100

resumen_na = pd.DataFrame({
    'Columna': df.columns,
    'Valores N/A': valores_na.values,
    'Porcentaje': porcentaje_na.values.round(2)
})

print(resumen_na[resumen_na['Valores N/A'] > 0])


# # ============================================================
# # 3. ESTRATEGIA DE LIMPIEZA POR COLUMNA
# # ============================================================

# print("\n" + "=" * 70)
# print("APLICANDO ESTRATEGIAS DE LIMPIEZA")
# print("=" * 70)

# # Pedido
# if 'pedido' in df.columns:
#     print("\n1. pedido:")
#     na_count = df['pedido'].isna().sum()
#     print(f"   - {na_count} valores N/A ({(na_count/len(df)*100):.2f}%)")
#     print("   - Decisión: Rellenar con 0 (sin pedido asociado)")
#     df['pedido'] = df['pedido'].fillna(0).astype(int)
#     print("   ✓ Completado")

# # Documento
# if 'Documento' in df.columns:
#     print("\n2. Documento:")
#     na_count = df['Documento'].isna().sum()
#     print(f"   - {na_count} valores N/A ({(na_count/len(df)*100):.2f}%)")
#     print("   - Decisión: Rellenar con 0 (documento no disponible)")
#     df['Documento'] = df['Documento'].fillna(0).astype(int)
#     print("   ✓ Completado")

# # Fecha Pedido
# if 'fecha_pedido' in df.columns and 'fecha_documento' in df.columns:
#     print("\n3. fecha_pedido:")
#     na_count = df['fecha_pedido'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     print("   - Decisión: Rellenar con fecha_documento cuando existe")
#     df['fecha_pedido'] = df['fecha_pedido'].fillna(df['fecha_documento'])
#     print("   ✓ Completado")

# # Fecha Documento
# if 'fecha_documento' in df.columns:
#     print("\n4. fecha_documento:")
#     na_count = df['fecha_documento'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     if na_count > 0:
#         print("   - Decisión: Rellenar con fecha mediana de la columna")
#         fecha_mediana = df['fecha_documento'].median()
#         df['fecha_documento'] = df['fecha_documento'].fillna(fecha_mediana)
#         # Solo rellenar fecha_pedido si existe
#         if 'fecha_pedido' in df.columns:
#             df['fecha_pedido'] = df['fecha_pedido'].fillna(fecha_mediana)
#         print(f"   ✓ Fecha de reemplazo: {fecha_mediana}")
#     else:
#         print("   ✓ No hay valores N/A")

# # Vendedor
# if 'Vendedor' in df.columns:
#     print("\n5. Vendedor:")
#     na_count = df['Vendedor'].isna().sum()
#     print(f"   - {na_count} valores N/A ({(na_count/len(df)*100):.2f}%)")
#     print("   - Decisión: Rellenar con 'SIN VENDEDOR'")
#     df['Vendedor'] = df['Vendedor'].fillna('SIN VENDEDOR')
#     print("   ✓ Completado")

# # CC (Centro de Costo)
# if 'CC' in df.columns:
#     print("\n6. CC (Centro de Costo):")
#     na_count = df['CC'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     print("   - Decisión: Rellenar con 0 (sin centro de costo)")
#     df['CC'] = df['CC'].fillna(0).astype(int)
#     print("   ✓ Completado")

# # Forma de Pago
# if 'F_Pag_Principal' in df.columns:
#     print("\n7. F_Pag_Principal (Forma de Pago):")
#     na_count = df['F_Pag_Principal'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     print("   - Decisión: Rellenar con 'NO ESPECIFICADO'")
#     df['F_Pag_Principal'] = df['F_Pag_Principal'].fillna('NO ESPECIFICADO')
#     print("   ✓ Completado")

# # Bruto (si existe)
# if 'Bruto' in df.columns:
#     print("\n8. Bruto:")
#     na_count = df['Bruto'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     if na_count > 0:
#         print("   - Decisión: Rellenar con 0")
#         df['Bruto'] = df['Bruto'].fillna(0)
#     print("   ✓ Completado")

# # Servicios (si existe)
# if 'Servicios' in df.columns:
#     print("\n9. Servicios:")
#     na_count = df['Servicios'].isna().sum()
#     print(f"   - {na_count} valores N/A")
#     if na_count > 0:
#         print("   - Decisión: Rellenar con 0")
#         df['Servicios'] = df['Servicios'].fillna(0)
#     print("   ✓ Completado")


# # ============================================================
# # 4. VERIFICACIÓN POST-LIMPIEZA
# # ============================================================

print("\n" + "=" * 70)
print("VERIFICACIÓN POST-LIMPIEZA")
print("=" * 70)

valores_na_final = df.isna().sum()
total_na_final = valores_na_final.sum()

print(f"\nTotal de valores N/A restantes: {total_na_final}")

if total_na_final == 0:
    print("✓ ¡ÉXITO! Todos los valores N/A han sido tratados")
else:
    print("\n⚠ Valores N/A restantes por columna:")
    print(valores_na_final[valores_na_final > 0])


# # ============================================================
# # 5. CREAR COLUMNAS ADICIONALES PARA ANÁLISIS
# # ============================================================

print("\n" + "=" * 70)
print("CREANDO COLUMNAS ADICIONALES PARA ANÁLISIS")
print("=" * 70)

# Extraer componentes de fecha
if 'fecha_documento' in df.columns:
    df['Año'] = df['fecha_documento'].dt.year
    df['Mes'] = df['fecha_documento'].dt.month
    df['Mes_Nombre'] = df['fecha_documento'].dt.month_name()
    df['Día'] = df['fecha_documento'].dt.day
    df['Día_Semana'] = df['fecha_documento'].dt.day_name()
    df['Trimestre'] = df['fecha_documento'].dt.quarter
    print("✓ Año, Mes, Día, Día_Semana, Trimestre")

# Clasificar tipo de documento
if 'Tipo_doc' in df.columns:
    df['Tipo_Doc_Categoria'] = df['Tipo_doc'].apply(
        lambda x: 'Factura' if 'FVE' in str(x) 
        else 'Boleta' if 'BOE' in str(x) or 'BSE' in str(x)
        else 'Otro'
    )
    print("✓ Tipo_Doc_Categoria")

# Rangos de ventas
if 'Total' in df.columns:
    df['Rango_Venta'] = pd.cut(df['Total'], 
                               bins=[-np.inf, 0, 50000, 200000, 500000, np.inf],
                               labels=['Negativo', 'Bajo', 'Medio', 'Alto', 'Muy Alto'])
    print("✓ Rango_Venta")


# # ============================================================
# # 6. ANÁLISIS EXPLORATORIO DE VARIABLES
# # ============================================================

# print("\n" + "=" * 70)
# print("ANÁLISIS EXPLORATORIO DE VARIABLES")
# print("=" * 70)

# # Identificar variables categóricas y numéricas
# lista_categorica = df.select_dtypes(include=['object', 'category']).columns.tolist()
# lista_numerica = df.select_dtypes(include=['int64', 'float64']).columns.tolist()

# print("\n📊 VARIABLES CATEGÓRICAS:")
# print(f"Total: {len(lista_categorica)} variables")
# for i, col in enumerate(lista_categorica, 1):
#     n_unique = df[col].nunique()
#     print(f"{i:2d}. {col:25s} - {n_unique:4d} valores únicos")

# print("\n📈 VARIABLES NUMÉRICAS:")
# print(f"Total: {len(lista_numerica)} variables")
# for i, col in enumerate(lista_numerica, 1):
#     min_val = df[col].min()
#     max_val = df[col].max()
#     mean_val = df[col].mean()
#     print(f"{i:2d}. {col:25s} - Min: {min_val:12,.0f} | Max: {max_val:12,.0f} | Prom: {mean_val:12,.0f}")

# # Análisis de variables categóricas con FOR
# print("\n" + "=" * 70)
# print("ANÁLISIS DETALLADO - VARIABLES CATEGÓRICAS")
# print("=" * 70)

# for col in lista_categorica:
#     print(f"\n{'=' * 70}")
#     print(f"Variable: {col}")
#     print('=' * 70)
    
#     # Frecuencias
#     freq = df[col].value_counts()
#     print(f"\nFrecuencias (top 10):")
#     for i, (val, count) in enumerate(freq.head(10).items(), 1):
#         porcentaje = (count / len(df)) * 100
#         print(f"{i:2d}. {str(val):40s} - {count:4d} ({porcentaje:5.2f}%)")
    
#     if len(freq) > 10:
#         print(f"    ... y {len(freq) - 10} valores más")

# # Análisis de variables numéricas con FOR
# print("\n" + "=" * 70)
# print("ANÁLISIS DETALLADO - VARIABLES NUMÉRICAS")
# print("=" * 70)

# for col in lista_numerica:
#     print(f"\n{'=' * 70}")
#     print(f"Variable: {col}")
#     print('=' * 70)
    
#     # Estadísticas descriptivas
#     stats = df[col].describe()
#     print(f"\nEstadísticas descriptivas:")
#     print(f"  Count:  {stats['count']:,.0f}")
#     print(f"  Mean:   {stats['mean']:,.2f}")
#     print(f"  Std:    {stats['std']:,.2f}")
#     print(f"  Min:    {stats['min']:,.2f}")
#     print(f"  25%:    {stats['25%']:,.2f}")
#     print(f"  50%:    {stats['50%']:,.2f}")
#     print(f"  75%:    {stats['75%']:,.2f}")
#     print(f"  Max:    {stats['max']:,.2f}")
    
#     # Valores nulos
#     na_count = df[col].isna().sum()
#     na_pct = (na_count / len(df)) * 100
#     print(f"\n  Valores N/A: {na_count} ({na_pct:.2f}%)")


# # ============================================================
# # 7. RESUMEN ESTADÍSTICO DE VENTAS
# # ============================================================

print("\n" + "=" * 70)
print("RESUMEN ESTADÍSTICO DE DATOS LIMPIOS")
print("=" * 70)

# Estadísticas generales
# columnas_ventas = [col for col in ['Neto', 'IVA', 'Bruto', 'Servicios', 'Total'] if col in df.columns]
columnas_ventas = [col for col in ['Neto', 'Bruto', 'Total'] if col in df.columns]
if columnas_ventas:
    print("\n📊 ESTADÍSTICAS GENERALES:")
    print(df[columnas_ventas].describe())

# Distribución temporal
if 'Mes_Nombre' in df.columns and 'Total' in df.columns:
    print("\n📅 DISTRIBUCIÓN TEMPORAL:")
    print(df.groupby('Mes_Nombre')['Total'].agg(['count', 'sum', 'mean']).round(2))

# Top clientes
if 'Cliente' in df.columns and 'Total' in df.columns:
    print("\n🏢 TOP 10 CLIENTES:")
    top_clientes = df.groupby('Cliente')['Total'].sum().sort_values(ascending=False).head(10)
    for i, (cliente, total) in enumerate(top_clientes.items(), 1):
        cliente_str = str(cliente)[:40]
        print(f"{i:2d}. {cliente_str:40s} ${total:,.0f}")

# Ventas por vendedor
if 'Vendedor' in df.columns and 'Total' in df.columns:
    print("\n👤 VENTAS POR VENDEDOR:")
    ventas_vendedor = df.groupby('Vendedor')['Total'].agg(['count', 'sum', 'mean']).sort_values('sum', ascending=False)
    print(ventas_vendedor.round(2))

# Análisis por tipo de documento
if 'Tipo_doc' in df.columns and 'Total' in df.columns:
    print("\n📄 VENTAS POR TIPO DE DOCUMENTO:")
    ventas_tipo = df.groupby('Tipo_doc')['Total'].agg(['count', 'sum', 'mean']).sort_values('sum', ascending=False)
    print(ventas_tipo.round(2))

# Análisis por forma de pago
if 'F_Pag_Principal' in df.columns and 'Total' in df.columns:
    print("\n💳 VENTAS POR FORMA DE PAGO:")
    ventas_pago = df.groupby('F_Pag_Principal')['Total'].agg(['count', 'sum', 'mean']).sort_values('sum', ascending=False)
    print(ventas_pago.round(2))


# # ============================================================
# # 8. GUARDAR DATOS LIMPIOS
# # ============================================================

# print("\n" + "=" * 70)
# print("GUARDANDO DATOS LIMPIOS")
# print("=" * 70)

# # Guardar en Excel
# archivo_limpio_xlsx = '/mnt/user-data/outputs/ventas2025_saratoga_limpio.xlsx'
# df.to_excel(archivo_limpio_xlsx, index=False, engine='openpyxl')
# print(f"✓ Excel guardado: {archivo_limpio_xlsx}")

# # Guardar en CSV
# archivo_limpio_csv = '/mnt/user-data/outputs/ventas2025_saratoga_limpio.csv'
# df.to_csv(archivo_limpio_csv, index=False, encoding='utf-8-sig')
# print(f"✓ CSV guardado: {archivo_limpio_csv}")


####################################### OTRO LIMPIEZA ################
# # Guardar reporte de limpieza
# reporte = f"""
# REPORTE DE LIMPIEZA DE DATOS - SARATOGA
# ========================================

# Archivo original: ventas2025.xlsx
# Fecha de limpieza: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

# DIMENSIONES:
# - Filas: {df.shape[0]}
# - Columnas originales: {df_original.shape[1]}
# - Columnas finales: {df.shape[1]}
# - Columnas agregadas: {df.shape[1] - df_original.shape[1]}

# VARIABLES CATEGÓRICAS: {len(lista_categorica)}
# {chr(10).join([f"  - {col}" for col in lista_categorica])}

# VARIABLES NUMÉRICAS: {len(lista_numerica)}
# {chr(10).join([f"  - {col}" for col in lista_numerica])}

# VALORES N/A TRATADOS:
# - pedido: Rellenado con 0
# - Documento: Rellenado con 0
# - fecha_pedido: Rellenado con fecha_documento
# - fecha_documento: Rellenado con mediana (si hay N/A)
# - Vendedor: Rellenado con 'SIN VENDEDOR'
# - CC: Rellenado con 0
# - F_Pag_Principal: Rellenado con 'NO ESPECIFICADO'
# - Bruto: Rellenado con 0 (si existe)
# - Servicios: Rellenado con 0 (si existe)

# COLUMNAS NUEVAS CREADAS:
# - Año, Mes, Mes_Nombre, Día, Día_Semana, Trimestre
# - Tipo_Doc_Categoria
# - Rango_Venta

# ESTADÍSTICAS CLAVE:
# - Total de ventas: ${df['Total'].sum() if 'Total' in df.columns else 0:,.0f}
# - Promedio por venta: ${df['Total'].mean() if 'Total' in df.columns else 0:,.0f}
# - Venta mínima: ${df['Total'].min() if 'Total' in df.columns else 0:,.0f}
# - Venta máxima: ${df['Total'].max() if 'Total' in df.columns else 0:,.0f}
# - Número de clientes únicos: {df['Cliente'].nunique() if 'Cliente' in df.columns else 0}
# - Número de vendedores: {df['Vendedor'].nunique() if 'Vendedor' in df.columns else 0}

# NOTAS:
# - Se encontraron {(df['Total'] < 0).sum() if 'Total' in df.columns else 0} ventas con valores negativos (devoluciones/NC)
# - Período de datos: {df['fecha_documento'].min() if 'fecha_documento' in df.columns else 'N/A'} a {df['fecha_documento'].max() if 'fecha_documento' in df.columns else 'N/A'}
# """

# archivo_reporte = '/mnt/user-data/outputs/reporte_limpieza_saratoga.txt'
# with open(archivo_reporte, 'w', encoding='utf-8') as f:
#     f.write(reporte)
# print(f"✓ Reporte guardado: {archivo_reporte}")

# print("\n" + "=" * 70)
# print("✓✓✓ PROCESO DE LIMPIEZA COMPLETADO EXITOSAMENTE ✓✓✓")
# print("=" * 70)
# print("\nArchivos generados:")
# print("1. ventas2025_saratoga_limpio.xlsx - Datos limpios en Excel")
# print("2. ventas2025_saratoga_limpio.csv - Datos limpios en CSV")
# print("3. reporte_limpieza_saratoga.txt - Reporte detallado")

###########################


"""
ANÁLISIS AVANZADO CON GROUPBY Y PIVOT
======================================
Análisis detallado por Tipo_Doc_Categoria y Trimestre

Autor: Puelche
Fecha: Diciembre 2024
"""



# ============================================================
# 1. ANÁLISIS CON GROUPBY - TIPO_DOC_CATEGORIA
# ============================================================

print("\n" + "=" * 80)
print("1. ANÁLISIS GROUPBY - TIPO DE DOCUMENTO")
print("=" * 80)

# 1.1 Agrupación básica por Tipo_Doc_Categoria
print("\n1.1 ESTADÍSTICAS POR TIPO DE DOCUMENTO:")
print("-" * 80)

grupo_tipo = df.groupby('Tipo_Doc_Categoria').agg({
    'Total': ['count', 'sum', 'mean', 'median', 'std', 'min', 'max'],
    'Neto': ['sum', 'mean'],
    'IVA': ['sum', 'mean']
})

# Renombrar columnas para mejor legibilidad
grupo_tipo.columns = ['_'.join(col).strip() for col in grupo_tipo.columns.values]
print(grupo_tipo.round(2))

# 1.2 Análisis detallado con múltiples funciones
print("\n1.2 ANÁLISIS DETALLADO POR TIPO DE DOCUMENTO:")
print("-" * 80)

for tipo in df['Tipo_Doc_Categoria'].unique():
    print(f"\n{'=' * 80}")
    print(f"TIPO: {tipo}")
    print('=' * 80)
    
    datos_tipo = df[df['Tipo_Doc_Categoria'] == tipo]
    
    print(f"\nCantidad de transacciones: {len(datos_tipo)}")
    print(f"Porcentaje del total:      {(len(datos_tipo)/len(df)*100):.2f}%")
    print(f"\nVentas:")
    print(f"  Total:    ${datos_tipo['Total'].sum():,.0f}")
    print(f"  Promedio: ${datos_tipo['Total'].mean():,.0f}")
    print(f"  Mediana:  ${datos_tipo['Total'].median():,.0f}")
    print(f"  Mínimo:   ${datos_tipo['Total'].min():,.0f}")
    print(f"  Máximo:   ${datos_tipo['Total'].max():,.0f}")
    print(f"  Desv.Est: ${datos_tipo['Total'].std():,.0f}")
    
    # Clientes únicos por tipo
    print(f"\nClientes únicos: {datos_tipo['Cliente'].nunique()}")
    
    # Top 3 clientes por tipo
    print("\nTop 3 Clientes:")
    top3 = datos_tipo.groupby('Cliente')['Total'].sum().nlargest(3)
    for i, (cliente, total) in enumerate(top3.items(), 1):
        print(f"  {i}. {cliente[:50]:50s} ${total:,.0f}")


# ============================================================
# 2. ANÁLISIS CON GROUPBY - TRIMESTRE
# ============================================================

print("\n" + "=" * 80)
print("2. ANÁLISIS GROUPBY - TRIMESTRE")
print("=" * 80)

# 2.1 Agrupación básica por Trimestre
print("\n2.1 ESTADÍSTICAS POR TRIMESTRE:")
print("-" * 80)

grupo_trim = df.groupby('Trimestre').agg({
    'Total': ['count', 'sum', 'mean', 'median', 'std', 'min', 'max'],
    'Neto': ['sum', 'mean'],
    'IVA': ['sum', 'mean']
})

grupo_trim.columns = ['_'.join(col).strip() for col in grupo_trim.columns.values]
print(grupo_trim.round(2))

# 2.2 Análisis detallado por trimestre
print("\n2.2 ANÁLISIS DETALLADO POR TRIMESTRE:")
print("-" * 80)

for trimestre in sorted(df['Trimestre'].unique()):
    print(f"\n{'=' * 80}")
    print(f"TRIMESTRE {trimestre}")
    print('=' * 80)
    
    datos_trim = df[df['Trimestre'] == trimestre]
    
    print(f"\nCantidad de transacciones: {len(datos_trim)}")
    print(f"Porcentaje del total:      {(len(datos_trim)/len(df)*100):.2f}%")
    print(f"\nVentas:")
    print(f"  Total:    ${datos_trim['Total'].sum():,.0f}")
    print(f"  Promedio: ${datos_trim['Total'].mean():,.0f}")
    print(f"  Mediana:  ${datos_trim['Total'].median():,.0f}")
    
    # Meses incluidos
    meses = datos_trim['Mes_Nombre'].unique()
    print(f"\nMeses incluidos: {', '.join(meses)}")
    
    # Distribución por mes
    print("\nVentas por mes:")
    ventas_mes = datos_trim.groupby('Mes_Nombre')['Total'].sum().sort_values(ascending=False)
    for mes, total in ventas_mes.items():
        print(f"  {mes:10s}: ${total:,.0f}")


# ============================================================
# 3. GROUPBY MÚLTIPLE - TIPO_DOC_CATEGORIA Y TRIMESTRE
# ============================================================

print("\n" + "=" * 80)
print("3. ANÁLISIS GROUPBY MÚLTIPLE - TIPO DOC + TRIMESTRE")
print("=" * 80)

# 3.1 Agrupación por ambas variables
print("\n3.1 VENTAS POR TIPO DE DOCUMENTO Y TRIMESTRE:")
print("-" * 80)

grupo_multiple = df.groupby(['Tipo_Doc_Categoria', 'Trimestre']).agg({
    'Total': ['count', 'sum', 'mean'],
    'Neto': 'sum',
    'IVA': 'sum'
})

grupo_multiple.columns = ['_'.join(col).strip() for col in grupo_multiple.columns.values]
print(grupo_multiple.round(2))

# 3.2 Porcentaje de participación
print("\n3.2 PARTICIPACIÓN PORCENTUAL:")
print("-" * 80)

participacion = df.groupby(['Tipo_Doc_Categoria', 'Trimestre'])['Total'].sum().unstack(fill_value=0)
participacion_pct = (participacion / participacion.sum().sum() * 100)
print("\nPorcentaje del total general:")
print(participacion_pct.round(2))

# 3.3 Análisis detallado combinado
print("\n3.3 ANÁLISIS DETALLADO COMBINADO:")
print("-" * 80)

for tipo in sorted(df['Tipo_Doc_Categoria'].unique()):
    print(f"\n{'=' * 80}")
    print(f"TIPO: {tipo}")
    print('=' * 80)
    
    for trimestre in sorted(df['Trimestre'].unique()):
        datos_filtrados = df[(df['Tipo_Doc_Categoria'] == tipo) & 
                             (df['Trimestre'] == trimestre)]
        
        if len(datos_filtrados) > 0:
            print(f"\n  Trimestre {trimestre}:")
            print(f"    Transacciones: {len(datos_filtrados):3d}")
            print(f"    Total:         ${datos_filtrados['Total'].sum():,.0f}")
            print(f"    Promedio:      ${datos_filtrados['Total'].mean():,.0f}")
        else:
            print(f"\n  Trimestre {trimestre}: Sin datos")


# ============================================================
# 4. TABLAS PIVOT - TIPO_DOC_CATEGORIA VS TRIMESTRE
# ============================================================

print("\n" + "=" * 80)
print("4. TABLAS PIVOT - TIPO DOC VS TRIMESTRE")
print("=" * 80)

# 4.1 Pivot: Total de ventas
print("\n4.1 PIVOT - TOTAL DE VENTAS ($):")
print("-" * 80)

pivot_total = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_total.round(0))

# 4.2 Pivot: Cantidad de transacciones
print("\n4.2 PIVOT - CANTIDAD DE TRANSACCIONES:")
print("-" * 80)

pivot_count = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='count',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_count)

# 4.3 Pivot: Ticket promedio
print("\n4.3 PIVOT - TICKET PROMEDIO ($):")
print("-" * 80)

pivot_mean = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='mean',
    fill_value=0
)
print(pivot_mean.round(0))

# 4.4 Pivot: Ventas Netas
print("\n4.4 PIVOT - VENTAS NETAS ($):")
print("-" * 80)

pivot_neto = pd.pivot_table(
    df,
    values='Neto',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_neto.round(0))

# 4.5 Pivot: IVA recaudado
print("\n4.5 PIVOT - IVA RECAUDADO ($):")
print("-" * 80)

pivot_iva = pd.pivot_table(
    df,
    values='IVA',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_iva.round(0))

# 4.6 Pivot con múltiples agregaciones
print("\n4.6 PIVOT - MÚLTIPLES MÉTRICAS:")
print("-" * 80)

pivot_multi = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc=['count', 'sum', 'mean', 'median'],
    fill_value=0
)
print(pivot_multi.round(0))


# ============================================================
# 5. ANÁLISIS DE TENDENCIAS Y CRECIMIENTO
# ============================================================

print("\n" + "=" * 80)
print("5. ANÁLISIS DE TENDENCIAS Y CRECIMIENTO")
print("=" * 80)

# 5.1 Crecimiento trimestral por tipo
print("\n5.1 CRECIMIENTO TRIMESTRAL POR TIPO DE DOCUMENTO:")
print("-" * 80)

for tipo in sorted(df['Tipo_Doc_Categoria'].unique()):
    print(f"\n{tipo}:")
    datos_tipo = df[df['Tipo_Doc_Categoria'] == tipo]
    ventas_trim = datos_tipo.groupby('Trimestre')['Total'].sum()
    
    print("  Trimestre  |  Ventas       |  Crecimiento")
    print("  " + "-" * 45)
    
    for i, (trim, venta) in enumerate(ventas_trim.items()):
        if i == 0:
            print(f"  Q{trim}         |  ${venta:>11,.0f}  |  -")
        else:
            trim_anterior = list(ventas_trim.index)[i-1]
            venta_anterior = ventas_trim[trim_anterior]
            crecimiento = ((venta - venta_anterior) / venta_anterior * 100) if venta_anterior != 0 else 0
            print(f"  Q{trim}         |  ${venta:>11,.0f}  |  {crecimiento:+.1f}%")

# 5.2 Share de mercado por trimestre
print("\n5.2 PARTICIPACIÓN DE MERCADO POR TRIMESTRE (%):")
print("-" * 80)

share_pivot = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Trimestre',
    aggfunc='sum',
    fill_value=0
)

# Calcular porcentaje por trimestre
share_pct = share_pivot.div(share_pivot.sum(axis=0), axis=1) * 100
print(share_pct.round(2))


# ============================================================
# 6. PIVOT CON OTRAS DIMENSIONES
# ============================================================

print("\n" + "=" * 80)
print("6. ANÁLISIS PIVOT ADICIONALES")
print("=" * 80)

# 6.1 Pivot: Tipo Doc vs Vendedor
print("\n6.1 PIVOT - TIPO DOC VS VENDEDOR (Total Ventas):")
print("-" * 80)

pivot_vendedor = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='Vendedor',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_vendedor.round(0))

# 6.2 Pivot: Tipo Doc vs Forma de Pago
print("\n6.2 PIVOT - TIPO DOC VS FORMA DE PAGO (Total Ventas):")
print("-" * 80)

pivot_pago = pd.pivot_table(
    df,
    values='Total',
    index='Tipo_Doc_Categoria',
    columns='F_Pag_Principal',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='TOTAL'
)
print(pivot_pago.round(0))

# 6.3 Pivot: Trimestre vs Mes
print("\n6.3 PIVOT - TRIMESTRE VS MES (Total Ventas):")
print("-" * 80)

pivot_mes = pd.pivot_table(
    df,
    values='Total',
    index='Trimestre',
    columns='Mes_Nombre',
    aggfunc='sum',
    fill_value=0
)
print(pivot_mes.round(0))


# ============================================================
# 7. GUARDAR RESULTADOS EN EXCEL
# ============================================================

print("\n" + "=" * 80)
print("7. GUARDANDO RESULTADOS EN EXCEL")
print("=" * 80)

# Crear archivo Excel con múltiples hojas
with pd.ExcelWriter('C:/Users/marce/OneDrive/Documentos/OneDrive/Documentos/Saratoga/analisis_groupby_pivot_saratoga.xlsx', 
                     engine='openpyxl') as writer:
    
    # Hoja 1: Resumen por Tipo Doc
    grupo_tipo.to_excel(writer, sheet_name='Resumen_TipoDoc')
    
    # Hoja 2: Resumen por Trimestre
    grupo_trim.to_excel(writer, sheet_name='Resumen_Trimestre')
    
    # Hoja 3: Pivot Total Ventas
    pivot_total.to_excel(writer, sheet_name='Pivot_TotalVentas')
    
    # Hoja 4: Pivot Cantidad
    pivot_count.to_excel(writer, sheet_name='Pivot_Cantidad')
    
    # Hoja 5: Pivot Ticket Promedio
    pivot_mean.to_excel(writer, sheet_name='Pivot_TicketPromedio')
    
    # Hoja 6: Pivot Neto
    pivot_neto.to_excel(writer, sheet_name='Pivot_Neto')
    
    # Hoja 7: Pivot IVA
    pivot_iva.to_excel(writer, sheet_name='Pivot_IVA')
    
    # Hoja 8: Share de Mercado
    share_pct.to_excel(writer, sheet_name='Share_Mercado')
    
    # Hoja 9: Tipo Doc vs Vendedor
    pivot_vendedor.to_excel(writer, sheet_name='TipoDoc_vs_Vendedor')
    
    # Hoja 10: Tipo Doc vs Forma Pago
    pivot_pago.to_excel(writer, sheet_name='TipoDoc_vs_FormaPago')

print("\n✓ Archivo guardado: analisis_groupby_pivot_saratoga.xlsx")
print("  Contiene 10 hojas con diferentes análisis")


# ============================================================
# 8. RESUMEN EJECUTIVO
# ============================================================

print("\n" + "=" * 80)
print("8. RESUMEN EJECUTIVO")
print("=" * 80)

print("\n📊 HALLAZGOS CLAVE:")

# Por Tipo de Documento
print("\n1. POR TIPO DE DOCUMENTO:")
tipo_summary = df.groupby('Tipo_Doc_Categoria')['Total'].agg(['count', 'sum', 'mean'])
tipo_summary = tipo_summary.sort_values('sum', ascending=False)
for tipo, datos in tipo_summary.iterrows():
    pct = (datos['sum'] / df['Total'].sum() * 100)
    print(f"   {tipo:10s}: {datos['count']:3.0f} trans | ${datos['sum']:>12,.0f} ({pct:5.1f}%) | Ticket: ${datos['mean']:>9,.0f}")

# Por Trimestre
print("\n2. POR TRIMESTRE:")
trim_summary = df.groupby('Trimestre')['Total'].agg(['count', 'sum', 'mean'])
for trim, datos in trim_summary.iterrows():
    pct = (datos['sum'] / df['Total'].sum() * 100)
    print(f"   Q{trim}: {datos['count']:3.0f} trans | ${datos['sum']:>12,.0f} ({pct:5.1f}%) | Ticket: ${datos['mean']:>9,.0f}")

# Mejor combinación
print("\n3. MEJOR COMBINACIÓN TIPO DOC + TRIMESTRE:")
mejor_combo = df.groupby(['Tipo_Doc_Categoria', 'Trimestre'])['Total'].sum().nlargest(5)
for i, ((tipo, trim), total) in enumerate(mejor_combo.items(), 1):
    print(f"   {i}. {tipo} - Q{trim}: ${total:,.0f}")

print("\n" + "=" * 80)
print("✓✓✓ ANÁLISIS GROUPBY Y PIVOT COMPLETADO ✓✓✓")
print("=" * 80)