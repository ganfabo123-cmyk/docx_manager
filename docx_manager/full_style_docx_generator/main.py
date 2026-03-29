from server import app

if __name__ == '__main__':
    """
    Server starter for the full style docx generator
    """
    app.run(host='0.0.0.0',debug=True, port=5000)