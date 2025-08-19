"""
Módulo principal para extraer datos de factura Air-e PDF y exportar a Excel
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
    pdf_path = "118860 - SDL AIR-E.pdf"  # Cambiar por la ruta de tu PDF
    output_excel = "factura_aire_datos.xlsx"
    
    print("=" * 60)
    print("EXTRACTOR DE DATOS DE FACTURA AIR-E")
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
        
        # Opcional: Guardar texto extraído para debugging
        if '--debug' in sys.argv:
            pdf_reader.guardar_texto_archivo("texto_extraido_aire.txt")
        
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
        
        # Opcional: También guardar como CSV
        if '--csv' in sys.argv:
            csv_file = "factura_aire_datos.csv"
            writer.guardar_como_csv(datos_extraidos, csv_file)
        
        # Mostrar estadísticas finales
        print("\n" + "=" * 60)
        print("PROCESO COMPLETADO EXITOSAMENTE")
        print("=" * 60)
        print(f"📁 Archivo de entrada: {pdf_path}")
        print(f"📊 Archivo de salida: {output_excel}")
        
        # Mostrar algunos datos clave extraídos
        if 'informacion_general' in datos_extraidos:
            info = datos_extraidos['informacion_general']
            if 'nic' in info:
                print(f"📋 NIC: {info['nic']}")
            if 'documento_equivalente' in info:
                print(f"📄 Documento Equivalente: {info['documento_equivalente']}")
            if 'total_pagar' in info:
                print(f"💰 Total: ${info['total_pagar']:,.0f}")
        
        if 'datos_cliente' in datos_extraidos:
            cliente = datos_extraidos['datos_cliente']
            if 'nombre' in cliente:
                print(f"👤 Cliente: {cliente['nombre']}")
            if 'nit' in cliente:
                print(f"🏢 NIT: {cliente['nit']}")
        
        if 'periodo_facturacion' in datos_extraidos:
            periodo = datos_extraidos['periodo_facturacion']
            if 'fecha_vencimiento' in periodo:
                print(f"📅 Vencimiento: {periodo['fecha_vencimiento']}")
        
        # Mostrar resumen detallado si se solicita
        if '--verbose' in sys.argv or '-v' in sys.argv:
            print("\n" + extractor.generar_resumen())
        
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
    
    dependencias_opcionales = {
        'pdfplumber': 'pdfplumber',
        'pypdf': 'pypdf'
    }
    
    faltantes = []
    for modulo, nombre_pip in dependencias.items():
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(nombre_pip)
    
    if faltantes:
        print("❌ Faltan las siguientes dependencias obligatorias:")
        for dep in faltantes:
            print(f"  • {dep}")
        print("\nPara instalarlas, ejecuta:")
        print(f"pip install {' '.join(faltantes)}")
        sys.exit(1)
    
    # Verificar dependencias opcionales
    opcionales_faltantes = []
    for modulo, nombre_pip in dependencias_opcionales.items():
        try:
            __import__(modulo)
        except ImportError:
            opcionales_faltantes.append(nombre_pip)
    
    if opcionales_faltantes:
        print("\n⚠️ Dependencias opcionales no instaladas (recomendadas):")
        for dep in opcionales_faltantes:
            print(f"  • {dep}")
        print("\nPara mejor rendimiento, instala con:")
        print(f"pip install {' '.join(opcionales_faltantes)}")
        print()


def mostrar_ayuda():
    """Mostrar información de ayuda"""
    print("""
Uso: python main.py [opciones]

Opciones:
  --help, -h       Mostrar esta ayuda
  --verbose, -v    Mostrar información detallada durante el proceso
  --debug          Guardar archivo de texto extraído para debugging
  --csv            También guardar los datos en formato CSV
  --pdf PATH       Especificar ruta del archivo PDF (por defecto: factura_aire.pdf)
  --output PATH    Especificar archivo de salida Excel (por defecto: factura_aire_datos.xlsx)

Ejemplos:
  python main.py
  python main.py --verbose --csv
  python main.py --pdf mi_factura.pdf --output resultado.xlsx
  python main.py --debug -v
""")


def procesar_argumentos():
    """Procesar argumentos de línea de comandos"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Extractor de datos de factura Air-e',
        add_help=False
    )
    
    parser.add_argument('--help', '-h', action='store_true', 
                       help='Mostrar ayuda')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Mostrar información detallada')
    parser.add_argument('--debug', action='store_true',
                       help='Guardar texto extraído para debugging')
    parser.add_argument('--csv', action='store_true',
                       help='También guardar en formato CSV')
    parser.add_argument('--pdf', type=str, default='factura_aire.pdf',
                       help='Ruta del archivo PDF')
    parser.add_argument('--output', type=str, default='factura_aire_datos.xlsx',
                       help='Archivo de salida Excel')
    
    args = parser.parse_args()
    
    if args.help:
        mostrar_ayuda()
        sys.exit(0)
    
    return args


if __name__ == "__main__":
    # Procesar argumentos de línea de comandos
    if len(sys.argv) > 1:
        if '--help' in sys.argv or '-h' in sys.argv:
            mostrar_ayuda()
            sys.exit(0)
        
        args = procesar_argumentos()
        
        # Actualizar configuración global con los argumentos
        if args.pdf:
            # Modificar el pdf_path en main()
            import __main__
            __main__.pdf_path = args.pdf
        if args.output:
            import __main__
            __main__.output_excel = args.output
    
    # Validar dependencias antes de ejecutar
    validar_dependencias()
    
    # Ejecutar programa principal
    main()