from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from flask_mysqldb import MySQL
import MySQLdb.cursors
import re
import os
from werkzeug.utils import secure_filename
from utils import handleMisProd, handleDailyProd, deleteConvertedXLS, sendForgotPasswordMail
from flask_cors import CORS
from flask_bcrypt import Bcrypt
from uuid import uuid4

app = Flask(__name__)
bcrypt = Bcrypt(app)
CORS(app)

app.secret_key = "fluid_mech"

app.config["MYSQL_HOST"] = "localhost"
app.config["MYSQL_USER"] = "root"
app.config["MYSQL_PASSWORD"] = "sarthak333"
app.config["MYSQL_DB"] = "fluid_control"

mysql = MySQL(app)


@app.route('/')
@app.route("/login", methods=["POST"])
def login():
    if request.method == 'POST':
        username = request.json.get("username")
        password = request.json.get("password")
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute(
            "SELECT * FROM accounts WHERE username=%s", (username,))
        account = cursor.fetchone()
        # print(account)
        # account = None
        # print(str())
        if account and bcrypt.check_password_hash(account['password'], password):
            # session["loggedin"] = True
            # session["id"] = account["id"]
            # session["username"] = account["username"]
            return jsonify({
                "message": "Loggen in Successfully"
            }), 200
        else:
            return jsonify({
                "message": "Incorrect username or password"
            }), 400
    else:
        return jsonify({
            "message": "Server Error!"
        }), 401


@app.route("/logout")
def logout():
    # session.pop("loggedin", None)
    # session.pop("id", None)
    # session.pop("username", None)
    return jsonify({
        "message": "Logged Out!"
    }), 200


@app.route('/resetpassword', methods=['POST'])
def resetPassword():
    uniqueID = request.json.get("uniqueid")
    newPassword = request.json.get("password")
    if uniqueID == None:
        return jsonify({"msg": "Please Provide UniqueID"}), 400
    elif newPassword == None:
        return jsonify({"msg": "Please Provide New Password"}), 400
    else:
        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute(
            "SELECT a.* FROM accounts a INNER JOIN forgotpassword fp ON a.id = fp.userid WHERE fp.uniqueid =%s", (uniqueID,))
        account = cursor.fetchone()
        if account:
            newPassword = bcrypt.generate_password_hash(
                newPassword).decode("utf-8")
            cursor.execute(
                "DELETE FROM forgotpassword WHERE uniqueid =%s", (uniqueID,))
            print(int(account['id']))
            cursor.execute(
                "UPDATE accounts SET password=%s WHERE id=%s", (newPassword, int(account['id']),))
            mysql.connection.commit()
            return jsonify({}), 200
        else:
            return jsonify({}), 400


@ app.route('/forgotpassword', methods=['POST'])
def forgotPassword():
    email = request.json.get("email")
    if (email == None):
        return jsonify({"msg": "Please Provide an valid Email"}), 400
    cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    cursor.execute(
        "SELECT * FROM accounts WHERE email=%s", (email,))
    account = cursor.fetchone()
    if account:
        uniqueID = uuid4()
        cursor.execute(
            'INSERT INTO forgotpassword VALUES (NULL, %s, %s )', (account['id'], uniqueID,))
        mysql.connection.commit()
        try:
            sendForgotPasswordMail(
                account['email'], account['username'], uniqueID)
            return jsonify({"msg": "Password Reset Link is Sent Successfully, Please Check your Email"}), 200
        except:
            return jsonify({}), 500
    else:
        return jsonify({"msg": "Please Provide an valid Email"}), 400


@ app.route('/register', methods=['GET', 'POST'])
def register():
    msg = ''
    if request.method == 'POST':
        username = request.json.get('username')
        password = request.json.get('password')
        email = request.json.get('email')
        firstname = request.json.get('firstname')
        lastname = request.json.get('lastname')
        date_of_birth = request.json.get('dob')

        password = bcrypt.generate_password_hash(password).decode("utf-8")

        cursor = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
        cursor.execute(
            'SELECT * FROM accounts WHERE username = % s', (username, ))
        account = cursor.fetchone()
        if account:
            return jsonify({
                "message": 'Account already exists !',
            }), 400
        elif not re.match(r'[^@]+@[^@]+\.[^@]+', email):
            return jsonify({
                "message": 'Invalid email address !',
            }), 400
        elif not re.match(r'[A-Za-z0-9]+', username):
            return jsonify({
                "message": 'Username must contain only characters and numbers !',
            }), 400
        elif not username or not password or not email:
            return jsonify({
                "message": 'Please fill out the form !',
            }), 400
        else:
            cursor.execute(
                'INSERT INTO accounts VALUES (NULL, %s, %s, %s,%s,%s,%s)', (firstname, lastname, username, password, email, date_of_birth))
            mysql.connection.commit()
            msg = 'You have successfully registered !'
            return jsonify({
                "message": msg,
            }), 200
    else:
        msg = 'Please fill out the form !'
        return jsonify({
            "message": msg,
        }), 401


def allowedFiles(fileName):
    if not "." in fileName:
        return False

    ext = fileName.rsplit(".", 1)[1]

    if ext.upper() in ["XLS", "XLSX"]:
        return True
    else:
        return False


@ app.route("/dailyprod", methods=["POST"])
def dailyProdHandler():
    if request.method == "POST":
        if request.files:
            f = request.files["file"]

            if f.filename == "":
                return jsonify({}), 401

            if f.filename.endswith("xlsx"):
                return jsonify({}), 500

            if not allowedFiles(f.filename):
                return jsonify({}), 401

            else:
                try:
                    fileName = secure_filename(f.filename)
                    f.save(fileName)
                    res = handleDailyProd(fileName)
                    deleteConvertedXLS(fileName)
                    return jsonify(res)
                except Exception:
                    print(Exception)
                    return jsonify({}), 500
        else:
            return jsonify({}), 401,
    else:
        return jsonify({}), 500


@ app.route("/misprod", methods=["POST"])
def misProdHandler():
    if request.method == "POST":
        if request.files:
            f = request.files["file"]

            if f.filename == "":
                return jsonify({}), 401

            if not allowedFiles(f.filename):
                return jsonify({}), 401

            else:
                fileName = secure_filename(f.filename)
                f.save(fileName)
                res = handleMisProd(fileName)
                deleteConvertedXLS(fileName)
                return jsonify(res)
        else:
            return jsonify({}), 401
    else:
        return jsonify({}), 500


if __name__ == "__main__":
    app.run(debug=True)
