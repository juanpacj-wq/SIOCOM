"""
Módulo para leer y extraer texto de archivos PDF
"""

import re
from typing import Optional, List
import sys

# Intentar múltiples librerías para mayor compatibilidad
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False


class PDFReader:
    """Clase para manejar la lectura de archivos PDF con múltiples métodos"""
    
    def __init__(self, pdf_path: str):
        """
        Inicializar el lector de PDF
        
        Args:
            pdf_path: Ruta al archivo PDF
        """
        self.pdf_path = pdf_path
        self.texto_completo = ""
        self.texto_por_pagina = []
        self.metodo_usado = None
    
    def extraer_texto(self) -> str:
        """
        Extraer todo el texto del PDF usando el mejor método disponible
        
        Returns:
            Texto completo del PDF como string
        """
        # Intentar con pdfplumber primero (más robusto)
        if PDFPLUMBER_AVAILABLE:
            print("  📚 Intentando con pdfplumber...")
            texto = self._extraer_con_pdfplumber()
            if texto:
                self.metodo_usado = "pdfplumber"
                return texto
        
        # Intentar con pypdf (versión más nueva)
        if PYPDF_AVAILABLE:
            print("  📚 Intentando con pypdf...")
            texto = self._extraer_con_pypdf()
            if texto:
                self.metodo_usado = "pypdf"
                return texto
        
        # Intentar con PyPDF2
        if PYPDF2_AVAILABLE:
            print("  📚 Intentando con PyPDF2...")
            texto = self._extraer_con_pypdf2()
            if texto:
                self.metodo_usado = "PyPDF2"
                return texto
        
        # Si no hay librerías disponibles
        if not any([PDFPLUMBER_AVAILABLE, PYPDF_AVAILABLE, PYPDF2_AVAILABLE]):
            print("❌ No se encontraron librerías PDF instaladas.")
            print("Por favor instala una de las siguientes:")
            print("  pip install pdfplumber")
            print("  pip install pypdf")
            print("  pip install PyPDF2")
            sys.exit(1)
        
        # Si ningún método funcionó
        raise Exception("No se pudo extraer texto del PDF con ningún método disponible")
    
    def _extraer_con_pdfplumber(self) -> Optional[str]:
        """Extraer texto usando pdfplumber"""
        try:
            import pdfplumber
            
            texto_completo = []
            with pdfplumber.open(self.pdf_path) as pdf:
                num_paginas = len(pdf.pages)
                print(f"  📄 Procesando {num_paginas} página(s) con pdfplumber...")
                
                for i, pagina in enumerate(pdf.pages):
                    texto_pagina = pagina.extract_text()
                    if texto_pagina:
                        texto_pagina = self._limpiar_texto(texto_pagina)
                        self.texto_por_pagina.append(texto_pagina)
                        texto_completo.append(texto_pagina)
                        print(f"  ✓ Página {i + 1} procesada")
                
                if texto_completo:
                    self.texto_completo = '\n'.join(texto_completo)
                    return self.texto_completo
                    
        except Exception as e:
            print(f"  ⚠️ pdfplumber falló: {str(e)[:100]}")
            return None
    
    def _extraer_con_pypdf(self) -> Optional[str]:
        """Extraer texto usando pypdf (versión moderna)"""
        try:
            from pypdf import PdfReader
            
            texto_completo = []
            with open(self.pdf_path, 'rb') as archivo:
                lector_pdf = PdfReader(archivo)
                num_paginas = len(lector_pdf.pages)
                print(f"  📄 Procesando {num_paginas} página(s) con pypdf...")
                
                for i in range(num_paginas):
                    pagina = lector_pdf.pages[i]
                    texto_pagina = pagina.extract_text()
                    if texto_pagina:
                        texto_pagina = self._limpiar_texto(texto_pagina)
                        self.texto_por_pagina.append(texto_pagina)
                        texto_completo.append(texto_pagina)
                        print(f"  ✓ Página {i + 1} procesada")
                
                if texto_completo:
                    self.texto_completo = '\n'.join(texto_completo)
                    return self.texto_completo
                    
        except Exception as e:
            print(f"  ⚠️ pypdf falló: {str(e)[:100]}")
            return None
    
    def _extraer_con_pypdf2(self) -> Optional[str]:
        """Extraer texto usando PyPDF2"""
        try:
            import PyPDF2
            
            texto_completo = []
            with open(self.pdf_path, 'rb') as archivo:
                lector_pdf = PyPDF2.PdfReader(archivo)
                num_paginas = len(lector_pdf.pages)
                print(f"  📄 Procesando {num_paginas} página(s) con PyPDF2...")
                
                for i in range(num_paginas):
                    try:
                        pagina = lector_pdf.pages[i]
                        texto_pagina = pagina.extract_text()
                        if texto_pagina:
                            texto_pagina = self._limpiar_texto(texto_pagina)
                            self.texto_por_pagina.append(texto_pagina)
                            texto_completo.append(texto_pagina)
                            print(f"  ✓ Página {i + 1} procesada")
                    except Exception as e:
                        print(f"  ⚠️ Error en página {i + 1}: {str(e)[:50]}")
                        # Continuar con la siguiente página
                        continue
                
                if texto_completo:
                    self.texto_completo = '\n'.join(texto_completo)
                    return self.texto_completo
                    
        except Exception as e:
            print(f"  ⚠️ PyPDF2 falló: {str(e)[:100]}")
            return None
    
    def _extraer_texto_alternativo(self) -> Optional[str]:
        """
        Método alternativo para extraer texto cuando los métodos principales fallan
        Usa OCR o técnicas alternativas
        """
        try:
            # Intentar con diferentes codificaciones
            import PyPDF2
            
            texto_completo = []
            with open(self.pdf_path, 'rb') as archivo:
                lector_pdf = PyPDF2.PdfReader(archivo)
                
                for i in range(len(lector_pdf.pages)):
                    try:
                        pagina = lector_pdf.pages[i]
                        # Intentar extracción básica sin procesamiento de fuentes
                        contenido = pagina.get_contents()
                        if contenido:
                            # Buscar texto directamente en el stream
                            texto_raw = str(contenido.get_data())
                            # Extraer texto visible con regex
                            matches = re.findall(r'\((.*?)\)', texto_raw)
                            texto_pagina = ' '.join(matches)
                            if texto_pagina:
                                texto_completo.append(texto_pagina)
                    except:
                        continue
                
                if texto_completo:
                    return '\n'.join(texto_completo)
                    
        except Exception as e:
            print(f"  ⚠️ Método alternativo falló: {str(e)[:100]}")
            return None
    
    def _limpiar_texto(self, texto: str) -> str:
        """
        Limpiar y normalizar el texto extraído del PDF
        
        Args:
            texto: Texto sin procesar del PDF
            
        Returns:
            Texto limpio y normalizado
        """
        if not texto:
            return ""
        
        # Eliminar caracteres especiales problemáticos pero mantener la estructura
        texto = texto.replace('\x00', '')
        texto = texto.replace('\ufeff', '')
        
        # Normalizar espacios múltiples pero mantener saltos de línea
        texto = re.sub(r'[ \t]+', ' ', texto)
        
        # Eliminar espacios al inicio y final de cada línea
        lineas = texto.split('\n')
        lineas = [linea.strip() for linea in lineas]
        texto = '\n'.join(lineas)
        
        # Eliminar líneas vacías múltiples
        texto = re.sub(r'\n{3,}', '\n\n', texto)
        
        return texto
    
    def obtener_texto_pagina(self, num_pagina: int) -> Optional[str]:
        """
        Obtener el texto de una página específica
        
        Args:
            num_pagina: Número de página (base 0)
            
        Returns:
            Texto de la página o None si no existe
        """
        if 0 <= num_pagina < len(self.texto_por_pagina):
            return self.texto_por_pagina[num_pagina]
        return None
    
    def buscar_patron(self, patron: str, flags=0) -> List[str]:
        """
        Buscar un patrón regex en todo el texto del PDF
        
        Args:
            patron: Expresión regular a buscar
            flags: Flags de regex (opcional)
            
        Returns:
            Lista de coincidencias encontradas
        """
        if not self.texto_completo:
            self.extraer_texto()
        
        matches = re.findall(patron, self.texto_completo, flags)
        return matches
    
    def extraer_seccion(self, inicio: str, fin: str) -> Optional[str]:
        """
        Extraer una sección del texto entre dos marcadores
        
        Args:
            inicio: Patrón de inicio de la sección
            fin: Patrón de fin de la sección
            
        Returns:
            Texto de la sección o None si no se encuentra
        """
        if not self.texto_completo:
            self.extraer_texto()
        
        patron = f"{re.escape(inicio)}(.*?){re.escape(fin)}"
        match = re.search(patron, self.texto_completo, re.DOTALL)
        
        if match:
            return match.group(1).strip()
        return None
    
    def obtener_estadisticas(self) -> dict:
        """
        Obtener estadísticas del documento PDF
        
        Returns:
            Diccionario con estadísticas del documento
        """
        if not self.texto_completo:
            self.extraer_texto()
        
        estadisticas = {
            'num_paginas': len(self.texto_por_pagina),
            'num_caracteres': len(self.texto_completo),
            'num_lineas': self.texto_completo.count('\n') + 1,
            'num_palabras': len(self.texto_completo.split()),
            'metodo_extraccion': self.metodo_usado
        }
        
        # Contar números (posibles valores monetarios)
        numeros = re.findall(r'\d+[\d,\.]*', self.texto_completo)
        estadisticas['num_valores_numericos'] = len(numeros)
        
        return estadisticas
    
    def guardar_texto_archivo(self, ruta_salida: str = "texto_extraido.txt"):
        """
        Guardar el texto extraído en un archivo de texto
        
        Args:
            ruta_salida: Ruta del archivo de salida
        """
        if not self.texto_completo:
            self.extraer_texto()
        
        try:
            with open(ruta_salida, 'w', encoding='utf-8') as archivo:
                archivo.write(self.texto_completo)
            print(f"  ✓ Texto guardado en: {ruta_salida}")
            print(f"  ℹ️ Método usado: {self.metodo_usado}")
        except Exception as e:
            print(f"  ❌ Error al guardar archivo: {e}")
            raise
    
    def extraer_texto_con_fallback(self) -> str:
        """
        Extraer texto con múltiples métodos de respaldo
        
        Returns:
            Texto extraído o cadena vacía si todos los métodos fallan
        """
        metodos = [
            ("pdfplumber", self._extraer_con_pdfplumber, PDFPLUMBER_AVAILABLE),
            ("pypdf", self._extraer_con_pypdf, PYPDF_AVAILABLE),
            ("PyPDF2", self._extraer_con_pypdf2, PYPDF2_AVAILABLE),
            ("alternativo", self._extraer_texto_alternativo, PYPDF2_AVAILABLE)
        ]
        
        for nombre, metodo, disponible in metodos:
            if disponible:
                print(f"  🔄 Intentando con método: {nombre}")
                try:
                    texto = metodo()
                    if texto:
                        self.metodo_usado = nombre
                        print(f"  ✅ Éxito con: {nombre}")
                        return texto
                except Exception as e:
                    print(f"  ⚠️ {nombre} falló: {str(e)[:50]}")
                    continue
        
        # Si todos fallan, retornar texto vacío
        print("  ⚠️ Advertencia: No se pudo extraer texto, usando texto vacío")
        return ""


# Funciones auxiliares para uso directo
def leer_pdf_rapido(pdf_path: str) -> str:
    """
    Función rápida para extraer texto de un PDF
    
    Args:
        pdf_path: Ruta al archivo PDF
        
    Returns:
        Texto completo del PDF
    """
    reader = PDFReader(pdf_path)
    return reader.extraer_texto()


def buscar_en_pdf(pdf_path: str, patron: str) -> List[str]:
    """
    Buscar un patrón en un PDF
    
    Args:
        pdf_path: Ruta al archivo PDF
        patron: Expresión regular a buscar
        
    Returns:
        Lista de coincidencias
    """
    reader = PDFReader(pdf_path)
    return reader.buscar_patron(patron)


def verificar_librerias_pdf():
    """Verificar qué librerías PDF están disponibles"""
    print("\n📚 Librerías PDF disponibles:")
    print(f"  • pdfplumber: {'✅' if PDFPLUMBER_AVAILABLE else '❌'}")
    print(f"  • pypdf: {'✅' if PYPDF_AVAILABLE else '❌'}")
    print(f"  • PyPDF2: {'✅' if PYPDF2_AVAILABLE else '❌'}")
    
    if not any([PDFPLUMBER_AVAILABLE, PYPDF_AVAILABLE, PYPDF2_AVAILABLE]):
        print("\n⚠️ No hay librerías PDF instaladas!")
        print("Instala al menos una con:")
        print("  pip install pdfplumber  # Recomendado")
        print("  pip install pypdf")
        print("  pip install PyPDF2")
        return False
    return True