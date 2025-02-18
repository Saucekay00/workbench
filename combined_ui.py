from flask import Flask, render_template, request

app = Flask(__name__)


# Route for rendering the form
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Handle the form submission
        uploaded_file = request.files.get('certFile')  # Adjust 'certFile' to match the name in your form
        if uploaded_file:
            return f"File '{uploaded_file.filename}' received successfully!"
        return "No file uploaded.", 400

    # Render the HTML form on GET
    return render_template('combined.html')


if __name__ == '__main__':
    app.run(debug=True, port=5003)
