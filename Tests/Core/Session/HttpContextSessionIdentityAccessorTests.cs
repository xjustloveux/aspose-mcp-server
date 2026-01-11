using AsposeMcpServer.Core.Session;
using Microsoft.AspNetCore.Http;
using Moq;

namespace AsposeMcpServer.Tests.Core.Session;

public class HttpContextSessionIdentityAccessorTests
{
    #region Constructor Tests

    [Fact]
    public void Constructor_WithHttpContextAccessor_ShouldNotThrow()
    {
        var mockAccessor = new Mock<IHttpContextAccessor>();

        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        Assert.NotNull(accessor);
    }

    #endregion

    #region GetCurrentIdentity Tests - No Context

    [Fact]
    public void GetCurrentIdentity_WithNullContext_ShouldReturnAnonymous()
    {
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns((HttpContext?)null);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.True(identity.IsAnonymous);
    }

    #endregion

    #region GetCurrentIdentity Tests - With Context Items

    [Fact]
    public void GetCurrentIdentity_WithGroupIdOnly_ShouldReturnIdentityWithGroupId()
    {
        var context = new DefaultHttpContext { Items = { ["GroupId"] = "group1" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Equal("group1", identity.GroupId);
        Assert.Null(identity.UserId);
    }

    [Fact]
    public void GetCurrentIdentity_WithUserIdOnly_ShouldReturnIdentityWithUserId()
    {
        var context = new DefaultHttpContext { Items = { ["UserId"] = "user1" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Null(identity.GroupId);
        Assert.Equal("user1", identity.UserId);
    }

    [Fact]
    public void GetCurrentIdentity_WithBothValues_ShouldReturnFullIdentity()
    {
        var context = new DefaultHttpContext { Items = { ["GroupId"] = "group1", ["UserId"] = "user1" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Equal("group1", identity.GroupId);
        Assert.Equal("user1", identity.UserId);
        Assert.False(identity.IsAnonymous);
    }

    [Fact]
    public void GetCurrentIdentity_WithNoItems_ShouldReturnIdentityWithNullValues()
    {
        var context = new DefaultHttpContext();
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Null(identity.GroupId);
        Assert.Null(identity.UserId);
    }

    [Fact]
    public void GetCurrentIdentity_WithSpecialCharacters_ShouldPreserveValues()
    {
        var context = new DefaultHttpContext
            { Items = { ["GroupId"] = "group:with:colons", ["UserId"] = "user@domain.com" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Equal("group:with:colons", identity.GroupId);
        Assert.Equal("user@domain.com", identity.UserId);
    }

    [Fact]
    public void GetCurrentIdentity_WithNonStringItems_ShouldConvertToString()
    {
        var context = new DefaultHttpContext
            { Items = { ["GroupId"] = 12345, ["UserId"] = new Guid("12345678-1234-1234-1234-123456789012") } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity = accessor.GetCurrentIdentity();

        Assert.Equal("12345", identity.GroupId);
        Assert.Equal("12345678-1234-1234-1234-123456789012", identity.UserId);
    }

    #endregion

    #region Multiple Calls Tests

    [Fact]
    public void GetCurrentIdentity_CalledMultipleTimes_ShouldReturnNewInstancesEachTime()
    {
        var context = new DefaultHttpContext { Items = { ["GroupId"] = "group1", ["UserId"] = "user1" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity1 = accessor.GetCurrentIdentity();
        var identity2 = accessor.GetCurrentIdentity();

        Assert.NotSame(identity1, identity2);
        Assert.Equal(identity1.GroupId, identity2.GroupId);
        Assert.Equal(identity1.UserId, identity2.UserId);
    }

    [Fact]
    public void GetCurrentIdentity_WithChangingContext_ShouldReflectChanges()
    {
        var context = new DefaultHttpContext { Items = { ["GroupId"] = "group1", ["UserId"] = "user1" } };
        var mockAccessor = new Mock<IHttpContextAccessor>();
        mockAccessor.Setup(x => x.HttpContext).Returns(context);
        var accessor = new HttpContextSessionIdentityAccessor(mockAccessor.Object);

        var identity1 = accessor.GetCurrentIdentity();
        Assert.Equal("group1", identity1.GroupId);

        context.Items["GroupId"] = "group2";
        var identity2 = accessor.GetCurrentIdentity();
        Assert.Equal("group2", identity2.GroupId);
    }

    #endregion
}
