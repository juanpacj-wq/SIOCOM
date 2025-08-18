"""
Módulo principal para extraer datos de factura ESSA PDF y exportar a Excel
"""

import sys
import os
from pathlib import Path
from pdf_reader import PDFReader
from data_extractor import DataExtractor
from excel_writer import ExcelWriter


def main():
    """Función principal del programa"""
    
    # Configuración
    pdf_path = "119325 - SDL ESSA.pdf"  # Cambiar por la ruta de tu PDF
    output_excel = "factura_essa_datos.xlsx"
    
    print("=" * 60)
    print("EXTRACTOR DE DATOS DE FACTURA ESSA")
    print("=" * 60)
    
    # Verificar que el archivo existe
    if not os.path.exists(pdf_path):
        print(f"❌ Error: No se encontró el archivo {pdf_path}")
        print("Por favor, asegúrate de que el archivo PDF esté en la carpeta correcta.")
        sys.exit(1)
    
    try:
        # Paso 1: Leer el PDF
        print("\n📄 Leyendo archivo PDF...")
        pdf_reader = PDFReader(pdf_path)
        texto_completo = pdf_reader.extraer_texto()
        
        if not texto_completo:
            print("❌ Error: No se pudo extraer texto del PDF")
            sys.exit(1)
        
        print(f"✅ Texto extraído exitosamente ({len(texto_completo)} caracteres)")
        
        # Paso 2: Extraer datos con regex
        print("\n🔍 Extrayendo datos con expresiones regulares...")
        extractor = DataExtractor(texto_completo)
        datos_extraidos = extractor.extraer_todos_los_datos()
        
        # Mostrar resumen de datos extraídos
        print(f"\n📊 Resumen de datos extraídos:")
        for categoria, valores in datos_extraidos.items():
            if valores:
                if isinstance(valores, dict):
                    print(f"  • {categoria}: {len(valores)} campos")
                elif isinstance(valores, list):
                    print(f"  • {categoria}: {len(valores)} items")
                else:
                    print(f"  • {categoria}: Extraído")
        
        # Paso 3: Guardar en Excel
        print(f"\n💾 Guardando datos en Excel: {output_excel}")
        writer = ExcelWriter(output_excel)
        writer.escribir_datos(datos_extraidos)
        
        print(f"✅ Archivo Excel creado exitosamente: {output_excel}")
        
        # Mostrar estadísticas finales
        print("\n" + "=" * 60)
        print("PROCESO COMPLETADO EXITOSAMENTE")
        print("=" * 60)
        print(f"📁 Archivo de entrada: {pdf_path}")
        print(f"📊 Archivo de salida: {output_excel}")
        
        # Mostrar algunos datos clave extraídos
        if 'informacion_general' in datos_extraidos:
            info = datos_extraidos['informacion_general']
            if 'numero_factura' in info:
                print(f"📄 Factura: {info['numero_factura']}")
            if 'total_pagar' in info:
                print(f"💰 Total: ${info['total_pagar']:,.0f}")
        
    except FileNotFoundError as e:
        print(f"❌ Error: Archivo no encontrado - {e}")
        sys.exit(1)
    except PermissionError as e:
        print(f"❌ Error: Sin permisos para acceder al archivo - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def validar_dependencias():
    """Validar que las dependencias necesarias estén instaladas"""
    dependencias = {
        'PyPDF2': 'PyPDF2',
        'openpyxl': 'openpyxl',
        'pandas': 'pandas'
    }
    
    faltantes = []
    for modulo, nombre_pip in dependencias.items():
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(nombre_pip)
    
    if faltantes:
        print("❌ Faltan las siguientes dependencias:")
        for dep in faltantes:
            print(f"  • {dep}")
        print("\nPara instalarlas, ejecuta:")
        print(f"pip install {' '.join(faltantes)}")
        sys.exit(1)


if __name__ == "__main__":
    # Validar dependencias antes de ejecutar
    validar_dependencias()
    
    # Ejecutar programa principal
    main()