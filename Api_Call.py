import requests
import os


def post_log(utility, emailadree, status, session):
    url = os.getenv("LOGGING_URL")


    data = {
        "utility": utility,
        "user": emailadree,
        "status": status,
        "sessionId": session,
        "tower": "PSP"
    }

    try:
        response = requests.post(url, json=data, timeout=10)  # Send POST request
        print("HTTP Status Code:", response.status_code)
        print("Response Text:", response.text)
    except requests.exceptions.Timeout:
        print("Error: Request timed out")
    except requests.exceptions.ConnectionError:
        print("Error: Could not connect to the server")
    except requests.exceptions.HTTPError as http_err:
        print("HTTP Error:", http_err)
    except requests.exceptions.RequestException as err:
        print("Request Error:", err)
