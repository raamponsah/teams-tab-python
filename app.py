from flask import Flask, render_template_string

app = Flask(__name__)

# HTML template for the Teams tab
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hello Teams Tab</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background-color: #f3f2f1;
        }
        .container {
            text-align: center;
            padding: 20px;
            background-color: #ffffff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        h1 {
            color: #6264a7;
        }
        p {
            color: #333;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Hello, Teams!</h1>
        <p>Welcome to your custom Microsoft Teams tab app.</p>
    </div>
    <script src="https://statics.teams.cdn.office.net/sdk/v2.1.0/js/MicrosoftTeams.min.js"></script>
    <script>
        microsoftTeams.app.initialize().then(() => {
            console.log("Teams SDK initialized");
        }).catch(err => {
            console.error("Failed to initialize Teams SDK:", err);
        });
    </script>
</body>
</html>
"""

@app.route('/helloWorldTab')
def hello_teams():
    return render_template_string(html_template)

if __name__ == '__main__':
    app.run(port=3007, debug=True)