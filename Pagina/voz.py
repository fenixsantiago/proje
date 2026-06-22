"""
Asistente de voz con ola visual animada
-----------------------------------------
Escucha continuamente por el micrófono. Cuando detecta la palabra
"preséntate" en lo que dijiste, responde hablando con una voz natural
(usando Edge-TTS, gratis y sin necesidad de API key).

La ventana muestra una ola animada que se mueve al hablar y
un indicador de micrófono cuando está escuchando.

Requisitos:
    pip install SpeechRecognition pyaudio edge-tts pygame numpy
"""

import asyncio
import speech_recognition as sr
import edge_tts
import pygame
import pygame.gfxdraw
import os
import tempfile
import threading
import time
import math
import numpy as np

# -----------------------------------------------------------------------
# CONFIGURACIÓN
# -----------------------------------------------------------------------

PALABRA_CLAVE = "presentate"
TEXTO_PRESENTACION = (
    "Bienvenidos a nuestro proyecto final de Procesamiento de Datos para Modelos de Inteligencia Artificial. "
    "Mi nombre es Siscon y seré su asistente durante esta presentación. "
    "En este proyecto ce presentara el sitio wed de siscon  ya que  se trata de un sitio interactivo que que ayudan a analizar contratos de aprendizaje, niveles de formación, empresas contratantes y riesgo de deserción. "
    "Esperamos que esta demostración sea de su interés. "
    
)

VOZ = "es-CO-GonzaloNeural"

# -----------------------------------------------------------------------
# ESTADO GLOBAL (compartido entre hilos)
# -----------------------------------------------------------------------

estado = {
    "hablando": False,       # True mientras el TTS reproduce audio
    "escuchando": False,     # True mientras el mic está abierto
    "texto_ui": "Escuchando...",
}

# -----------------------------------------------------------------------
# VENTANA PYGAME - OLA ANIMADA
# -----------------------------------------------------------------------

ANCHO, ALTO = 700, 260
FPS = 60

# Paleta
COLOR_FONDO   = (10, 10, 20)
COLOR_OLA_1   = (0, 200, 255)     # cian brillante
COLOR_OLA_2   = (80, 80, 255)     # azul medio
COLOR_OLA_3   = (0, 255, 180)     # verde agua (sutil)
COLOR_TEXTO   = (200, 220, 255)
COLOR_MIC_ON  = (0, 220, 120)
COLOR_MIC_OFF = (60, 60, 80)
COLOR_HABLA   = (0, 200, 255)


def dibujar_ola(superficie, tiempo, amplitud, color, fase=0.0, grosor=2, alpha=255):
    """
    Dibuja una ola senoidal suave sobre la superficie.
    amplitud: qué tan alta es la ola (0 = plana, 60 = enorme)
    """
    puntos = []
    n = ANCHO + 2
    for x in range(n):
        angulo = (x / ANCHO) * 2 * math.pi * 3 + tiempo + fase
        y_ola = math.sin(angulo) * amplitud
        # segunda armónica para hacerla menos aburrida
        y_ola += math.sin(angulo * 1.7 + fase) * amplitud * 0.35
        y = int(ALTO // 2 + y_ola)
        puntos.append((x, y))

    if len(puntos) > 1:
        # Creamos una superficie temporal para el alpha
        tmp = pygame.Surface((ANCHO, ALTO), pygame.SRCALPHA)
        pygame.draw.lines(tmp, (*color, alpha), False, puntos, grosor)
        superficie.blit(tmp, (0, 0))


def dibujar_indicador_mic(superficie, activo):
    """Pequeño círculo de micrófono en la esquina superior derecha."""
    cx, cy, r = ANCHO - 28, 28, 10
    color = COLOR_MIC_ON if activo else COLOR_MIC_OFF
    pygame.gfxdraw.filled_circle(superficie, cx, cy, r, (*color, 200))
    pygame.gfxdraw.aacircle(superficie, cx, cy, r, color)
    # símbolo de micrófono (rectángulo pequeño + arco)
    rect = pygame.Rect(cx - 4, cy - 7, 8, 10)
    pygame.draw.rect(superficie, COLOR_FONDO, rect, border_radius=4)
    pygame.draw.rect(superficie, color, rect, width=2, border_radius=4)
    pygame.draw.line(superficie, color, (cx, cy + 3), (cx, cy + 7), 2)
    pygame.draw.line(superficie, color, (cx - 5, cy + 7), (cx + 5, cy + 7), 2)


def loop_visual():
    """Hilo principal de pygame — ventana con la ola animada."""
    pygame.init()
    pantalla = pygame.display.set_mode((ANCHO, ALTO))
    pygame.display.set_caption("Asistente de Voz")
    reloj = pygame.time.Clock()

    try:
        fuente_grande = pygame.font.SysFont("segoeui", 22, bold=False)
        fuente_chica  = pygame.font.SysFont("segoeui", 14)
    except Exception:
        fuente_grande = pygame.font.Font(None, 24)
        fuente_chica  = pygame.font.Font(None, 16)

    t = 0.0          # tiempo continuo para animar la ola
    amp_suave = 0.0  # amplitud suavizada (interpolamos hacia el objetivo)

    corriendo = True
    while corriendo:
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                corriendo = False
            if evento.type == pygame.KEYDOWN and evento.key == pygame.K_ESCAPE:
                corriendo = False

        # ------------------------------------------------------------------
        # Determinar amplitud objetivo según el estado
        # ------------------------------------------------------------------
        if estado["hablando"]:
            amp_objetivo = 55.0
            velocidad    = 0.055   # más rápida cuando habla
        elif estado["escuchando"]:
            amp_objetivo = 14.0
            velocidad    = 0.030
        else:
            amp_objetivo = 4.0
            velocidad    = 0.025

        # Suavizar la amplitud (lerp)
        amp_suave += (amp_objetivo - amp_suave) * 0.08

        # ------------------------------------------------------------------
        # Fondo con gradiente vertical sutil
        # ------------------------------------------------------------------
        pantalla.fill(COLOR_FONDO)

        # Línea central de referencia (muy tenue)
        pygame.draw.line(pantalla, (30, 30, 50), (0, ALTO // 2), (ANCHO, ALTO // 2), 1)

        # ------------------------------------------------------------------
        # Dibujar 3 olas con distinto desfase, grosor y alpha
        # ------------------------------------------------------------------
        dibujar_ola(pantalla, t * 1.0,  amp_suave * 1.0,  COLOR_OLA_1, fase=0.0,  grosor=3, alpha=230)
        dibujar_ola(pantalla, t * 0.75, amp_suave * 0.65, COLOR_OLA_2, fase=1.1,  grosor=2, alpha=160)
        dibujar_ola(pantalla, t * 1.3,  amp_suave * 0.40, COLOR_OLA_3, fase=2.3,  grosor=1, alpha=100)

        # ------------------------------------------------------------------
        # Texto de estado
        # ------------------------------------------------------------------
        msg = estado["texto_ui"]
        surf_txt = fuente_grande.render(msg, True, COLOR_TEXTO)
        pantalla.blit(surf_txt, (20, ALTO - 40))

        # Etiqueta "HABLANDO" / "ESCUCHANDO"
        if estado["hablando"]:
            etiqueta = fuente_chica.render("● HABLANDO", True, COLOR_HABLA)
        else:
            etiqueta = fuente_chica.render("● ESCUCHANDO", True, COLOR_MIC_ON if estado["escuchando"] else COLOR_MIC_OFF)
        pantalla.blit(etiqueta, (20, 14))

        # Indicador micrófono
        dibujar_indicador_mic(pantalla, estado["escuchando"] and not estado["hablando"])

        pygame.display.flip()

        # Velocidad de avance de la ola
        t += velocidad
        reloj.tick(FPS)

    pygame.quit()


# -----------------------------------------------------------------------
# TEXTO A VOZ
# -----------------------------------------------------------------------

async def _generar_audio(texto: str, ruta: str):
    comunicador = edge_tts.Communicate(texto, VOZ)
    await comunicador.save(ruta)


def hablar(texto: str):
    print(f"[Asistente]: {texto}")
    estado["texto_ui"] = texto[:55] + ("..." if len(texto) > 55 else "")

    ruta_temp = os.path.join(tempfile.gettempdir(), "respuesta_asistente.mp3")
    asyncio.run(_generar_audio(texto, ruta_temp))

    # Usamos un mixer separado para no interferir con pygame principal
    mixer = pygame.mixer
    mixer.init()
    mixer.music.load(ruta_temp)

    estado["hablando"] = True
    mixer.music.play()

    while mixer.music.get_busy():
        time.sleep(0.05)

    estado["hablando"] = False
    estado["texto_ui"] = "Escuchando..."
    mixer.quit()


# -----------------------------------------------------------------------
# ESCUCHA DEL MICRÓFONO (corre en hilo aparte)
# -----------------------------------------------------------------------

def hilo_escucha():
    reconocedor = sr.Recognizer()
    microfono   = sr.Microphone()

    with microfono as fuente:
        print("Ajustando al ruido ambiente, un momento...")
        reconocedor.adjust_for_ambient_noise(fuente, duration=2)

    print("Listo. Di 'preséntate' para activar al asistente. Esc para cerrar.\n")

    while True:
        try:
            estado["escuchando"] = True
            with microfono as fuente:
                audio = reconocedor.listen(fuente, phrase_time_limit=5)
            estado["escuchando"] = False

            texto = reconocedor.recognize_google(audio, language="es-ES")
            texto_norm = (
                texto.lower()
                .replace("é", "e").replace("á", "a")
                .replace("í", "i").replace("ó", "o").replace("ú", "u")
            )

            print(f"[Escuchado]: {texto}")
            estado["texto_ui"] = f'"{texto}"'

            if PALABRA_CLAVE in texto_norm:
                hablar(TEXTO_PRESENTACION)

        except sr.UnknownValueError:
            pass
        except sr.RequestError as e:
            print(f"Error reconocimiento: {e}")
        except Exception as e:
            if "exit" in str(e).lower():
                break


# -----------------------------------------------------------------------
# PUNTO DE ENTRADA
# -----------------------------------------------------------------------

if __name__ == "__main__":
    # Lanzar el hilo de escucha en segundo plano
    t = threading.Thread(target=hilo_escucha, daemon=True)
    t.start()

    # La ventana visual corre en el hilo principal (pygame lo requiere)
    loop_visual()

    print("\n¡Hasta luego!")