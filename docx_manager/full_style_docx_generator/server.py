from flask import Flask
import os
from api.routes import register_routes

app = Flask(__name__)

# Configure upload folder
app.config['UPLOAD_FOLDER'] = 'temp'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Register API routes
register_routes(app)

if __name__ == '__main__':
    app.run(debug=True, port=5000)