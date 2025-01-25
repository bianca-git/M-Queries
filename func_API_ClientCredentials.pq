let
    //
    // 1) Define your OAuth token endpoint, client ID, client secret, scope, etc.
    //
    tokenUrl = "https://idp.example.com/oauth2/token", 
    clientId = "YOUR_CLIENT_ID",
    clientSecret = "YOUR_CLIENT_SECRET",
    scope = "https://api.example.com/.default",  // Replace with correct scope(s)

    //
    // 2) Construct the body for the 'client_credentials' token request.
    //    Check your IDP’s docs for required parameters (e.g., “grant_type”, “scope”).
    //
    requestBody = "grant_type=client_credentials" & 
                  "&client_id=" & clientId & 
                  "&client_secret=" & clientSecret &
                  "&scope=" & scope,

    //
    // 3) Obtain the JSON response from the token endpoint.
    //    The IDP will typically return an “access_token” field.
    //    Some providers also require “Content-Type: application/x-www-form-urlencoded”.
    //
    tokenResponse = 
        Json.Document(
            Web.Contents(
                tokenUrl,
                [
                    Headers = [
                        #"Content-Type" = "application/x-www-form-urlencoded",
                        Accept = "application/json"
                    ],
                    Content = Text.ToBinary(requestBody)
                ]
            )
        ),

    //
    // 4) Extract the access token from the JSON response. 
    //    Most IDPs return it under "access_token".
    //
    accessToken = tokenResponse[access_token],

    //
    // 5) Construct the authorization header for the subsequent API call.
    //    Format is typically "Bearer <access_token>".
    //
    authHeader = "Bearer " & accessToken,

    //
    // 6) Call the protected API endpoint using the retrieved access token.
    //    Adjust the “apiUrl” and “Headers” based on your API’s requirements.
    //
    apiUrl = "https://api.example.com/v1/resource",
    apiResponse = 
        Web.Contents(
            apiUrl,
            [
                Headers = [
                    Authorization = authHeader,
                    Accept = "application/json"
                ]
            ]
        ),

    //
    // 7) Parse the response from the protected resource (assuming JSON).
    //
    finalData = Json.Document(apiResponse)
in
    //
    // 8) Return the parsed data.
    //
    finalData
