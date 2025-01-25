let
    // 1) Define your authentication key/token (e.g., Bearer token).
    //    In practice, you might retrieve or generate this dynamically.
    AuthKey = "YOUR_AUTH_KEY",

    // 2) Define the API base URL or endpoint.
    //    Change this to match the exact endpoint you need.
    ApiUrl = "https://api.example.com/v1/data",

    // 3) Pass the AuthKey in the 'Authorization' header.
    //    Adjust the header format based on your API's requirements 
    //    (e.g., "Bearer <token>", "Token <key>", or custom schemes).
    ApiResponse = Web.Contents(
        ApiUrl,
        [
            Headers = [
                Authorization = "Bearer " & AuthKey,
                Accept = "application/json"
            ]
        ]
    ),

    // 4) Parse the JSON response (assuming the API returns JSON).
    JsonData = Json.Document(ApiResponse)
in
    // 5) Return the parsed data as your final query output.
    JsonData
