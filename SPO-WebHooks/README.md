## Troubleshooting Guide

### Certificate Issues
- Ensure that the certificate being used is valid and not expired.
- Check that the certificate's thumbprint is correctly specified in your application settings.
- Confirm that the certificate is installed on the server and correctly referenced in code.

### Azure AD Authentication Failures
- Verify that the application is registered correctly in Azure AD.
- Ensure that the required API permissions are granted and consented.
- Check the client ID and client secret for any typos or incorrect values.

### SharePoint Webhook Subscription Issues
- Ensure the notification URL is accessible and publicly reachable.
- Confirm that the correct resource URL is being used for the subscription.
- Check for any issues in the SharePoint admin center related to the site or subscription.

### Notification Reception Problems
- Verify that your endpoint is set up to handle notifications properly.
- Check your logs to see if notifications are being sent but not processed.
- Ensure that your firewall or security settings do not block incoming requests.

### Deployment Issues
- Check the deployment logs for any errors during the deployment process.
- Confirm that all environment variables are set correctly for production.
- Ensure that any necessary services are started and running.

### Debugging Tips
- Use logging to capture details about the application’s state when issues occur.
- Utilize Azure Application Insights to track exceptions and performance metrics.
- Test changes in a local or staging environment before production deployment.

### Error Reference Table
| Error Code | Description                              | Solution                                           |
|------------|------------------------------------------|------------------------------------------------------|
| 1001       | Invalid certificate                      | Check certificate validity and thumbprint.         |
| 2002       | Authentication failed                    | Verify Azure AD client ID and secret.               |
| 3003       | Failed to receive notification           | Check the notification URL and security settings.   |