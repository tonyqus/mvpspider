using Ganss.Excel;
using Microsoft.Playwright;
using Microsoft.Playwright.NUnit;
using Microsoft.VisualStudio.TestPlatform.CoreUtilities.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace MVPSpider;

[Parallelizable(ParallelScope.Self)]
[TestFixture]
public class MVPPageTests : PlaywrightTest
{
    private IAPIRequestContext APIRequest = null;
    private IAPIRequestContext FrontRequest = null;
    [SetUp]
    public async Task SetUpAPITesting()
    {
        await CreateAPIRequestContext();
    }

    private async Task CreateAPIRequestContext()
    {
        var headers = new Dictionary<string, string>();
        // We set this header per GitHub guidelines.
        headers.Add("Accept", "application/*.*");
        headers.Add("Content-Type", "application/json");

        APIRequest = await this.Playwright.APIRequest.NewContextAsync(new()
        {
            // All requests we send go to this API endpoint.
            BaseURL = "https://mavenapi-prod.azurewebsites.net",
            ExtraHTTPHeaders = headers,
        });

        FrontRequest = await this.Playwright.APIRequest.NewContextAsync(new()
        {
            // All requests we send go to this API endpoint.
            BaseURL = "https://mvp.microsoft.com"
        });
    }
    [Test]
    public async Task GetMVPSinglePageContent_LiangLiu()
    {
        var request = await this.APIRequest.GetAsync("/api/mvp/UserProfiles/public/98a7fa7d-6c60-eb11-a812-000d3a8ccaf5");
        Assert.True(request.Ok);
        var jsonEle = await request.JsonAsync();
        var node = jsonEle?.Get("userProfile");
        Assert.That(node?.Get("tenants")?.EnumerateArray().ToList()[0].GetString(), Is.EqualTo("MVP"));
        //get location
        Assert.That(node?.Get("addressCountryOrRegionName")?.GetString(), Is.EqualTo("China"));
        //get mvp category and year
        Assert.That(node?.Get("yearsInProgram")?.GetInt32(), Is.EqualTo(3));
        Assert.That(node?.Get("awardCategory")?.EnumerateArray().ToList()[0].GetString(), Is.EqualTo("Developer Technologies"));
        Assert.That(node?.Get("technologyFocusArea")?.EnumerateArray().ToList()[0].GetString(), Is.EqualTo(".NET"));

        //get biography
        var bio = node?.Get("biography")?.GetString();
        Assert.IsTrue(bio.Trim().StartsWith("Great passion for software development"));
        //recent activities
        //var activities = await Page.Locator(".raListTable tr").AllAsync();
        //Assert.IsTrue(activities.Count > 0);
    }
    [Test]
    public async Task GetMVPSinglePageContent_CyrusWong()
    {
        //https://mavenapi-prod.azurewebsites.net/api/Contributions/HighImpact/86da86ff-8786-ed11-aad1-000d3a197333/MVP
        //https://mavenapi-prod.azurewebsites.net/api/Events/HighImpact/86da86ff-8786-ed11-aad1-000d3a197333/MVP
        //https://mvp.microsoft.com/Locales/Activities.json
        //https://mvp.microsoft.com/Locales/UserProfile.json
        //https://mvp.microsoft.com/Locales/CountryOrRegion.json
        var request = await this.APIRequest.GetAsync("/api/mvp/UserProfiles/public/86da86ff-8786-ed11-aad1-000d3a197333");
        Assert.True(request.Ok);
        var jsonEle = await request.JsonAsync();
        var node = jsonEle?.Get("userProfile");
        //get location
        Assert.That(node?.Get("addressCountryOrRegionName")?.GetString(), Is.EqualTo("Hong Kong SAR"));
        //get mvp category and year
        Assert.That(node?.Get("yearsInProgram")?.GetInt32(), Is.EqualTo(1));
        //get contact and language info

        var links= node?.Get("userProfileSocialNetwork")?.EnumerateArray().ToList();
        Assert.That(links?.Count, Is.EqualTo(7));
        
        Assert.That(links?[0].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://www.linkedin.com/in/cyruswong"));
        Assert.That(links?[1].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://www.youtube.com/channel/UCjzFlDS8Zu8sIRJWeldfJ1w"));
        Assert.That(links?[2].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://techcommunity.microsoft.com/t5/user/viewprofilepage/user-id/1135201#profile"));
        Assert.That(links?[3].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://twitter.com/wongcyrus"));
        Assert.That(links?[4].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://github.com/wongcyrus"));
        Assert.That(links?[5].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://www.facebook.com/cywong.vtc"));
        Assert.That(links?[6].Get("socialNetworkImageLink")?.GetString(), Is.EqualTo("https://www.youtube.com/@CyrusWong"));
    }

    [Test]
    public async Task GetMVPSinglePageContent_JamesYeung()
    {
        var request = await this.APIRequest.GetAsync("/api/mvp/UserProfiles/public/8ad149c0-5c01-eb11-a815-000d3a8ccaf5");
        Assert.True(request.Ok);
        var jsonEle = await request.JsonAsync();
        var node = jsonEle?.Get("userProfile");
        Assert.That(node?.Get("firstName")?.GetString(), Is.EqualTo("James"));
        Assert.That(node?.Get("lastName")?.GetString(), Is.EqualTo("Yeung"));

        Assert.That(node?.Get("profilePictureUrl")?.GetString(), Is.EqualTo("https://mavenstorageprod.blob.core.windows.net/profile-pictures/8ad149c0-5c01-eb11-a815-000d3a8ccaf5"));

        //get mvp category and year
        Assert.That(node?.Get("awardCategory")?.EnumerateArray().ToList()[0].GetString(), Is.EqualTo("Developer Technologies"));
        var techFocus = node?.Get("technologyFocusArea")?.EnumerateArray().Select(x=> x.GetString()).ToList();
        Assert.That(techFocus?[0], Is.EqualTo("DevOps"));
        Assert.That(techFocus?[1], Is.EqualTo(".NET"));
        //get location
        Assert.That(node?.Get("addressCountryOrRegionName")?.GetString(), Is.EqualTo("China"));
        //get mvp category and year
        Assert.That(node?.Get("yearsInProgram")?.GetInt32(), Is.EqualTo(3));
        //get language info
        var langs=node?.Get("languages")?.EnumerateArray().Select(x => x.GetString()).ToList();
        Assert.That(langs?[0], Is.EqualTo("CHINESE_SIMPLIFIED_LANGUAGE"));
        Assert.That(langs?[1], Is.EqualTo("ENGLISH_LANGUAGE"));

        //get biography
        var bio = node?.Get("biography")?.GetString()?.Trim();
        Assert.IsTrue(bio?.StartsWith("I am the author of Ant Design Blazor"));

        //get contact and language info
        var links = node?.Get("userProfileSocialNetwork")?.EnumerateArray().ToList();
        Assert.That(links?.Count, Is.EqualTo(1));
        //recent activities
        //var activities = await Page.Locator(".raListTable tr").AllAsync();
        //Assert.IsTrue(activities.Count > 0);
    }
    [TearDown]
    public async Task TearDownAPITesting()
    {
        await APIRequest.DisposeAsync();
    }
}
