// Alternative: Device Code Flow for headless environments
// Replace the token acquisition section in CreateListWithCSOM.cs with this:

// authResult = await app.AcquireTokenWithDeviceCode(scopes, callback =>
// {
//     Console.WriteLine(callback.Message);
//     return Task.CompletedTask;
// }).ExecuteAsync();