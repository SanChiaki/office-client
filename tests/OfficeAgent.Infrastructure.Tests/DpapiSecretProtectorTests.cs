using OfficeAgent.Infrastructure.Security;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class DpapiSecretProtectorTests
    {
        [Fact]
        public void RoundTripsSecrets()
        {
            var protector = new DpapiSecretProtector();

            var protectedValue = protector.Protect("demo-secret");
            var plainText = protector.Unprotect(protectedValue);

            Assert.NotEqual("demo-secret", protectedValue);
            Assert.Equal("demo-secret", plainText);
        }
    }
}
