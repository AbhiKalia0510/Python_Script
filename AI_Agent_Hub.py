# -*- coding: utf-8 -*-
"""
Created on Tue Mar 17 14:52:22 2026

@author: abhishek.a
"""

# ==============================
# 🚀 AGENTHUB SaaS (SINGLE FILE)
# ==============================

from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
import uvicorn

app = FastAPI()

# =========================
# 🔌 COURIER API LAYER
# =========================
class CourierAPI:

    def check_all(self, pincode):
        return [
            {"name": "Delhivery", "serviceable": True, "cod": True, "prepaid": True, "tat": "1-2"},
            {"name": "Shiprocket", "serviceable": True, "cod": True, "prepaid": True, "tat": "2-3"},
            {"name": "BlueDart", "serviceable": True, "cod": True, "prepaid": True, "tat": "Next Day"},
            {"name": "DTDC", "serviceable": True, "cod": False, "prepaid": True, "tat": "2-3"},
            {"name": "Ecom Express", "serviceable": False, "cod": False, "prepaid": False, "tat": "-"},
            {"name": "XpressBees", "serviceable": True, "cod": True, "prepaid": True, "tat": "2-3"},
        ]

api = CourierAPI()

# =========================
# 🧠 AI ENGINES
# =========================
class CostEngine:
    def get_cost(self, name):
        return {
            "Delhivery": 45,
            "Shiprocket": 50,
            "BlueDart": 70,
            "DTDC": 55,
            "XpressBees": 48
        }.get(name, 60)


class RTOEngine:
    def predict(self, courier, cod):
        risk = 10
        if cod:
            risk += 20
        if courier == "Ecom Express":
            risk += 30
        return f"{risk}%"


class RecommendationEngine:
    def best(self, data):
        best = None
        score = 999

        for d in data:
            if not d["serviceable"]:
                continue

            s = 0
            s += 1 if d["tat"] == "Next Day" else 3
            s += 5 if not d["cod"] else 0

            if s < score:
                score = s
                best = d["name"]

        return best


cost_engine = CostEngine()
rto_engine = RTOEngine()
reco_engine = RecommendationEngine()


# =========================
# 📦 REQUEST MODEL
# =========================
class PincodeRequest(BaseModel):
    pincode: str


# =========================
# 🔗 API ENDPOINT
# =========================
@app.post("/check")
def check(req: PincodeRequest):
    data = api.check_all(req.pincode)

    for d in data:
        d["cost"] = cost_engine.get_cost(d["name"])
        d["rto"] = rto_engine.predict(d["name"], d["cod"])

    best = reco_engine.best(data)

    return {"data": data, "best": best}


# =========================
# 🌐 FRONTEND (INLINE UI)
# =========================
@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
    <head>
        <title>AgentHub SaaS</title>
        <style>
            body { background:#0f172a; color:white; font-family:sans-serif; padding:30px;}
            input { padding:10px; width:200px;}
            button { padding:10px; background:green; color:white; border:none;}
            table { width:100%; margin-top:20px; border-collapse:collapse;}
            th, td { padding:10px; border-bottom:1px solid #333;}
            .best { color:#22c55e; font-weight:bold;}
        </style>
    </head>
    <body>

        <h1>🚀 AgentHub Logistics AI</h1>

        <input id="pin" placeholder="Enter Pincode"/>
        <button onclick="check()">Check</button>

        <div id="output"></div>

        <script>
        async function check() {
            const pincode = document.getElementById("pin").value;

            const res = await fetch("/check", {
                method:"POST",
                headers:{"Content-Type":"application/json"},
                body: JSON.stringify({pincode})
            });

            const data = await res.json();

            let html = "<table><tr><th>Courier</th><th>Status</th><th>COD</th><th>TAT</th><th>Cost</th><th>RTO</th></tr>";

            data.data.forEach(c => {
                let best = c.name === data.best ? "best" : "";
                html += `<tr class="${best}">
                    <td>${c.name}</td>
                    <td>${c.serviceable ? "Active" : "Inactive"}</td>
                    <td>${c.cod ? "Yes" : "No"}</td>
                    <td>${c.tat}</td>
                    <td>₹${c.cost}</td>
                    <td>${c.rto}</td>
                </tr>`;
            });

            html += "</table>";

            html += `<h3>⭐ Recommended: ${data.best}</h3>`;

            document.getElementById("output").innerHTML = html;
        }
        </script>

    </body>
    </html>
    """


# =========================
# ▶️ RUN APP
# =========================
if __name__ == "__main__":
    uvicorn.run("agenthub_saas:app", host="0.0.0.0", port=8000, reload=True)