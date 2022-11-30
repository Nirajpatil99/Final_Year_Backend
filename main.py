from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from flask_mysqldb import MySQL
import MySQLdb.cursors
import re
import os
from werkzeug.utils import secure_filename
from utils import handleMisProd
from utils import deleteConvertedXLS
from utils import handleDailyProd

app = Flask(__name__)

app.secret_key = "fluid_mech"

app.config["MYSQL_HOST"] = "localhost"
app.config["MYSQL_USER"] = "root"
app.config["MYSQL_PASSWORD"] = ""
app.config["MYSQL_DB"] = "Fluid Logins"

mysql = MySQL(app)


@app.route('/')
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == 'POST':
        username = request.json.get("username")
        password = request.json.get("password")
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute(
            "SELECT * FROM accounts WHERE username = % s AND password = % s", (username, password, ))
        account = cursor.fetchone()

        if account:
            session["loggedin"] = True
            session["id"] = account["id"]
            session["username"] = account["username"]
            return jsonify({
                "status": 200,
                "message": "Loggen in Successfully"
            })
        else:
            return jsonify({
                "status": 0,
                "message": "Incorrect username or password"
            })
    else:
        return jsonify({
            "status": 0,
            "message": "Server Error!"
        })


@app.route("/logout")
def logout():
    session.pop("loggedin", None)
    session.pop("id", None)
    session.pop("username", None)
    return jsonify({
        "status": 200,
        "message": "Logged Out!"
    })


@app.route('/register', methods=['GET', 'POST'])
def register():
    msg = ''
    if request.method == 'POST':
        username = request.json.get('username')
        password = request.json.get('password')
        email = request.json.get('email')
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute(
            'SELECT * FROM accounts WHERE username = % s', (username, ))
        account = cursor.fetchone()
        if account:
            return jsonify({
                "status": 0,
                "message": 'Account already exists !',
            })
        elif not re.match(r'[^@]+@[^@]+\.[^@]+', email):
            return jsonify({
                "status": 0,
                "message": 'Invalid email address !',
            })
        elif not re.match(r'[A-Za-z0-9]+', username):
            return jsonify({
                "status": 0,
                "message": 'Username must contain only characters and numbers !',
            })
        elif not username or not password or not email:
            return jsonify({
                "status": 0,
                "message": 'Please fill out the form !',
            })
        else:
            cursor.execute(
                'INSERT INTO accounts VALUES (NULL, % s, % s, % s)', (username, password, email, ))
            mysql.connection.commit()
            msg = 'You have successfully registered !'
            return jsonify({
                "status": 200,
                "message": msg,
            })
    else:
        msg = 'Please fill out the form !'
        return jsonify({
            "status": 0,
            "message": msg,
        })


def allowedFiles(fileName):
    if not "." in fileName:
        return False

    ext = fileName.rsplit(".", 1)[1]

    if ext.upper() in ["XLS", "XLSX"]:
        return True
    else:
        return False


@app.route("/dailyProd", methods=["POST"])
def dailyProdHandler():
    if request.method == "POST":
        if request.files:
            f = request.files["file"]

            if f.filename == "":
                return jsonify({
                    "status": 0,
                    "message": "No file selected",
                })

            if not allowedFiles(f.filename):
                return jsonify({
                    "status": 0,
                    "message": "only excel files are allowed",
                })

            else:
                fileName = secure_filename(f.filename)
                f.save(fileName)
                res = handleDailyProd(fileName)
                deleteConvertedXLS(fileName)
                return jsonify({
                    "status": 200,
                    "message": "Data converted successfully!",
                    "response": res,
                    "mimetype": "application/json"
                })
        else:
            return jsonify({
                "status": 0,
                "message": "Failed to upload file!"
            })
    else:
        return jsonify({
            "status": 500,
            "message": "Server Error!"
        })

@app.route("/misprod", methods=["POST"])
def misProdHandler():
    if request.method == "POST":
        if request.files:
            f = request.files["file"]

            if f.filename == "":
                return jsonify({
                    "status": 0,
                    "message": "No file selected",
                })

            if not allowedFiles(f.filename):
                return jsonify({
                    "status": 0,
                    "message": "only excel files are allowed",
                })

            else:
                fileName = secure_filename(f.filename)
                f.save(fileName)
                res = handleMisProd(fileName)
                deleteConvertedXLS(fileName)
                return jsonify({
                    "status": 200,
                    "message": "Data converted successfully!",
                    "response": res,
                    "mimetype": "application/json"
                })
        else:
            return jsonify({
                "status": 0,
                "message": "Failed to upload file!"
            })
    else:
        return jsonify({
            "status": 500,
            "message": "Server Error!"
        })


if __name__ == "__main__":
    app.run(debug=True)
