"""Prepara el AndroidManifest generado por Capacitor para el escáner de códigos.

Se ejecuta dentro del workflow de GitHub Actions, justo después de `cap add android`.
Capacitor ya sabe pedir el permiso de cámara al WebView cuando la página llama a
getUserMedia(); aquí solo declaramos el permiso en el manifest.
"""
from pathlib import Path

manifest = Path("android/app/src/main/AndroidManifest.xml")
texto = manifest.read_text(encoding="utf-8")

if "android.permission.CAMERA" not in texto:
    texto = texto.replace(
        "<application",
        '    <uses-permission android:name="android.permission.CAMERA" />\n'
        '    <uses-feature android:name="android.hardware.camera" android:required="false" />\n\n'
        "    <application",
        1,
    )

# Capacitor sirve los archivos desde http://localhost dentro del WebView.
if "usesCleartextTraffic" not in texto:
    texto = texto.replace(
        "<application",
        '<application android:usesCleartextTraffic="true"',
        1,
    )

manifest.write_text(texto, encoding="utf-8")
print("AndroidManifest listo: cámara habilitada para el escáner")
