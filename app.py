from flask import Flask, request, jsonify
import pandas as pd
import openpyxl

app = Flask(__name__)

@app.route("/")
def home():
    return "Welcome to Baseball Simulation API"

@app.route("/simulate", methods=["POST"])
def simulate():
    try:
        # アップロードされたファイルを保存
        batter_file = request.files["batter_file"]
        pitcher_file = request.files["pitcher_file"]
        schedule_file = request.files["schedule_file"]

        batter_file.save("Y3.xlsx")
        pitcher_file.save("P5.xlsx")
        schedule_file.save("日程.xlsx")

        # あなたのシミュレーション関数を呼び出す
        from main import main
        main()

        return jsonify({"message": "シミュレーションが完了しました！結果がresults.xlsxに保存されました。"})
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == "__main__":
    app.run(debug=True)
