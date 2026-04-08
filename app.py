from flask import Flask, request, jsonify
import requests
import os

app = Flask(__name__)

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/generate", methods=["POST"])
def generate():
    try:
        # 1. Obtener los datos enviados por n8n
        data = request.json
        
        # Si n8n envía una lista (array) en lugar de un objeto, tomamos el primer elemento
        if isinstance(data, list):
            data = data[0]

        # 2. Configurar los encabezados para la API de Anthropic
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        
        if not api_key:
            return jsonify({"error": "La variable de entorno ANTHROPIC_API_KEY no está configurada en el servidor."}), 500

        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        }
        
        # 3. Preparar el payload para Claude
        # Usamos los datos que vienen de n8n o valores por defecto
        payload = {
            "model": data.get("model", "claude-3-5-sonnet-20240620"),
            "max_tokens": data.get("max_tokens", 8192),
            "messages": data.get("messages", [])
        }

        # 4. Hacer la petición POST a la API de Anthropic
        anthropic_url = "https://api.anthropic.com/v1/messages"
        response = requests.post(anthropic_url, headers=headers, json=payload)
        
        # 5. Manejar posibles errores de la API de Anthropic
        if response.status_code != 200:
             print(f"Error de Anthropic API: {response.status_code} - {response.text}")
             return jsonify({
                 "error": "Error al comunicarse con Claude",
                 "details": response.json()
             }), response.status_code

        # 6. Extraer y devolver la respuesta de Claude a n8n
        response_data = response.json()
        return jsonify(response_data)

    except Exception as e:
        print(f"Error interno del servidor: {str(e)}")
        return jsonify({"error": "Error interno del servidor", "details": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
