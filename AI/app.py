from flask import Flask, request, jsonify
import joblib

app = Flask(__name__)

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

model = joblib.load(os.path.join(BASE_DIR, "svm_model.pkl"))
vectorizer = joblib.load(os.path.join(BASE_DIR, "tfidf_vectorizer.pkl"))


@app.route("/predict", methods=["POST"])
def predict():

    data = request.get_json()

    text = data.get("text", "")

    vector = vectorizer.transform([text])

    prediction = model.predict(vector)[0]

    if int(prediction) == 1:
        result = "accepted"
    else:
        result = "rejected"

    return jsonify({
        "prediction": int(prediction),
        "result": result
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)