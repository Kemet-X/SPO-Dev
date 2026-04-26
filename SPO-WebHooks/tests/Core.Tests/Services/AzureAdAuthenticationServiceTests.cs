using Microsoft.Extensions.Logging;
using Moq;
using SPO.Webhooks.Core.Services;
using System.Security.Cryptography.X509Certificates;
using Xunit;

namespace SPO.Webhooks.Core.Tests.Services;

public class AzureAdAuthenticationServiceTests
{
    private readonly Mock<ILogger<AzureAdAuthenticationService>> _mockLogger;
    private readonly string _tenantId = "test-tenant-id";
    private readonly string _clientId = "test-client-id";

    public AzureAdAuthenticationServiceTests()
    {
        _mockLogger = new Mock<ILogger<AzureAdAuthenticationService>>();
    }

    [Fact]
    public void Constructor_WithValidParameters_ShouldSucceed()
    {
        // Arrange
        using var certificate = new X509Certificate2();

        // Act & Assert
        var service = new AzureAdAuthenticationService(
            _tenantId,
            _clientId,
            certificate,
            _mockLogger.Object);

        Assert.NotNull(service);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Constructor_WithNullOrEmptyTenantId_ShouldThrowArgumentException(string tenantId)
    {
        // Arrange
        using var certificate = new X509Certificate2();

        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
            new AzureAdAuthenticationService(tenantId, _clientId, certificate, _mockLogger.Object));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Constructor_WithNullOrEmptyClientId_ShouldThrowArgumentException(string clientId)
    {
        // Arrange
        using var certificate = new X509Certificate2();

        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
            new AzureAdAuthenticationService(_tenantId, clientId, certificate, _mockLogger.Object));
    }

    [Fact]
    public void Constructor_WithNullCertificate_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() =>
            new AzureAdAuthenticationService(_tenantId, _clientId, null!, _mockLogger.Object));
    }

    [Fact]
    public void GetAuthorityUrl_ShouldReturnCorrectUrl()
    {
        // Arrange
        using var certificate = new X509Certificate2();
        var service = new AzureAdAuthenticationService(
            _tenantId,
            _clientId,
            certificate,
            _mockLogger.Object);

        // Act
        var authorityUrl = service.GetAuthorityUrl();

        // Assert
        Assert.Equal($"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token", authorityUrl);
    }
}