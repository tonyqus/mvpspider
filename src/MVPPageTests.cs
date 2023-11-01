using Ganss.Excel;
using Microsoft.Playwright.NUnit;
using Microsoft.VisualStudio.TestPlatform.CoreUtilities.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MVPSpider;

[Parallelizable(ParallelScope.Self)]
[TestFixture]
public class MVPPageTests : PageTest
{
    [Test]
    public async Task GetMVPSinglePageContent_LiangLiu()
    {
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/PublicProfile/5004128?fullName=Liang%20Liu");
        await Expect(Page).ToHaveTitleAsync(new Regex("Liang"));
        //get mvp category and year
        var content2 = await Page.Locator(".ktbFdU .css-124").AllTextContentsAsync();
        //get location
        Assert.That(content2[1], Is.EqualTo("China"));
        Assert.That(content2[3], Is.EqualTo("3 years in the program"));

        var content = await Page.Locator(".sc-dlfnbm.bcaJjD .sc-gsTCUz.bhdLno").AllTextContentsAsync();

        Assert.That(content[5], Is.EqualTo("Developer Technologies"));
        Assert.That(content[7], Is.EqualTo(".NET"));

        //get biography
        var content3 = await Page.Locator(".sc-dlfnbm.sc-jSgupP.bcaJjD span").AllTextContentsAsync();
        var bio = content3[1];
        Assert.IsTrue(bio.Trim().StartsWith("Great passion for software development"));
        //recent activities
        //var activities = await Page.Locator(".raListTable tr").AllAsync();
        //Assert.IsTrue(activities.Count > 0);
    }
    [Test]
    public async Task GetMVPSinglePageContent_CyrusWong()
    {
        //https://mavenapi-prod.azurewebsites.net/api/mvp/UserProfiles/public/86da86ff-8786-ed11-aad1-000d3a197333
        //https://mavenapi-prod.azurewebsites.net/api/Contributions/HighImpact/86da86ff-8786-ed11-aad1-000d3a197333/MVP
        //https://mavenapi-prod.azurewebsites.net/api/Events/HighImpact/86da86ff-8786-ed11-aad1-000d3a197333/MVP
        //https://mvp.microsoft.com/Locales/Activities.json
        //https://mvp.microsoft.com/Locales/UserProfile.json
        //https://mvp.microsoft.com/Locales/CountryOrRegion.json

        await Page.GotoAsync("https://mvp.microsoft.com/en-US/MVP/profile/86da86ff-8786-ed11-aad1-000d3a197333");
        await Task.Delay(500);
        //get location
        var content2 = await Page.Locator(".ms-Stack").AllTextContentsAsync();

        Assert.That(content2[1], Is.EqualTo("Hong Kong SAR"));
        Assert.That(content2[3], Is.EqualTo("1 year in the program"));
        //get contact and language info
        var socialLinks = await Page.Locator(".onlineIdentityRow .otherContent a").AllAsync();
        var links = new List<string>();
        foreach (var link in socialLinks)
        {
            links.Add(await link.GetAttributeAsync("href"));
        }
        var result= String.Join("|", links.ToArray());
        Assert.IsTrue(links.Count==4);
        Assert.That(links[0], Is.EqualTo("https://www.facebook.com/victang0228"));
        Assert.That(links[1], Is.EqualTo("https://cn.linkedin.com/in/鑫岳-汤-568b31104"));
        Assert.That(links[2], Is.EqualTo("https://github.com/VicDynamics"));
        Assert.That(links[3], Is.EqualTo("https://blog.csdn.net/vic0228"));
    }
    [Test]
    public async Task GetMVPSinglePageContent_JamesYeung()
    {
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/PublicProfile/5003987?fullName=James%20%20Yeung");
        await Expect(Page).ToHaveTitleAsync(new Regex("James Yeung"));

        var images = await Page.Locator(".ms-Persona-image img").AllAsync();
        var imgUrl = await images[0].GetAttributeAsync("src");
        var name = await images[0].GetAttributeAsync("alt");
        Assert.IsNotEmpty(imgUrl);
        Assert.That(name, Is.EqualTo("James Yeung"));

        await Page.GetByLabel("Read more").ClickAsync();
        //get mvp category and year
        var content = await Page.Locator(".sc-dlfnbm.bcaJjD .sc-gsTCUz.bhdLno").AllTextContentsAsync();

        Assert.That(content[5], Is.EqualTo("Developer Technologies"));
        Assert.That(content[7], Is.EqualTo("GitHub & Azure DevOps, .NET"));
        //get language info
        Assert.That(content[9], Is.EqualTo("Chinese (Simplified), English"));

        //get location
        var content2 = await Page.Locator(".ktbFdU .css-124").AllTextContentsAsync();

        Assert.That(content2[1], Is.EqualTo("China"));
        Assert.That(content2[3], Is.EqualTo("3 years in the program"));
        //var state = await Page.Locator(".state").TextContentAsync();
        //Assert.That(state, Is.EqualTo("Shanghai"));

        //get biography
        var content3 = await Page.Locator(".sc-dlfnbm.sc-jSgupP.bcaJjD span").AllTextContentsAsync();
        var bio = content3[1];
        Assert.IsTrue(bio.Trim().StartsWith("I am the author of Ant Design Blazor"));

        //get contact and language info
        var buttons = await Page.Locator(".sc-dacFzL.fIUiIW button img").AllAsync();
        var otherInfo = new Dictionary<string, string>();
        foreach (var info in buttons)
        {
            var link = await info.GetAttributeAsync("title");
            var mediaName = await info.GetAttributeAsync("alt");
            otherInfo.Add(mediaName.Trim(), link.Trim());
        }
        Assert.That(otherInfo.Count, Is.EqualTo(1));
        //recent activities
        //var activities = await Page.Locator(".raListTable tr").AllAsync();
        //Assert.IsTrue(activities.Count > 0);
    }
    [Test]
    public async Task GetMVPDetailListAndSaveToExcel_Taiwan()
    {
        var UrlList = new List<string>();
        var mvpDetails = new List<MVPDetail>();

        for (int i = 1; i <= 3; i++)
        {
            var url = $"https://mvp.microsoft.com/en-us/MvpSearch?lo=Taiwan&sc=e&pn={i}";
            await GetUrlsFromOneUrl(url, UrlList);
        }
        var categories = new Dictionary<string, CategoryDetail>();
        foreach (var url in UrlList)
        {
            var MVPDetail = await VisitOnePage(url, categories);
            mvpDetails.Add(MVPDetail);
        }
        new ExcelMapper().Save("mvp_taiwan.xlsx", mvpDetails, "MVPs");
    }
    [Test]
    public async Task GetMVPDetailListAndSaveToExcel_ChinaMainland()
    {
        var UrlList = new List<string>();
        var mvpDetails = new List<MVPDetail>();

        for (int i = 1; i <= 8; i++)
        {
            var url = $"https://mvp.microsoft.com/en-us/MvpSearch?lo=China&sc=e&pn={i}";
            await GetUrlsFromOneUrl(url, UrlList);
        }
        var categories = new Dictionary<string, CategoryDetail>();
        foreach (var url in UrlList)
        {
            var MVPDetail = await VisitOnePage(url, categories);
            mvpDetails.Add(MVPDetail);
        }
        new ExcelMapper().Save("mvp_china.xlsx", mvpDetails, "MVPs");
    }
    [Test]
    public async Task GetChinaMVPUrls()
    {
        var visitedPages = new List<string>();
        var UrlList = new List<string>();

        for (int i = 1; i <= 8; i++)
        {
            visitedPages.Add($"https://mvp.microsoft.com/en-us/MvpSearch?lo=China&sc=e&pn={i}");
        }
        foreach (var url in visitedPages)
        {
            await GetUrlsFromOneUrl(url, UrlList);
        }
        Assert.IsTrue(UrlList.Count == 138);
    }
    private async Task GetUrlsFromOneUrl(string pageUrl, List<string> UrlList)
    {
        await Page.GotoAsync(pageUrl);
        var links = await Page.Locator(".profileListItem a").AllAsync();
        foreach (var link in links)
        {
            var url = await link.GetAttributeAsync("href");
            if(!UrlList.Contains(url))
                UrlList.Add(url);
        }
    }

    [Test]
    public async Task GetCategoriesForAllMVP_China()
    {
        var UrlList = new List<string>();

        for (int i = 1; i <= 8; i++)
        {
            var url = $"https://mvp.microsoft.com/en-us/MvpSearch?lo=China&sc=e&pn={i}";
            await GetUrlsFromOneUrl(url, UrlList);
        }
        var categories = new Dictionary<string, CategoryDetail>();
        foreach (var url in UrlList)
        {
            await VisitOnePage(url, categories);
        }
        Assert.That(categories.Count, Is.EqualTo(14));

        new ExcelMapper().Save("categories_china.xlsx", categories.Values, "MVP Categories");
    }
    private async Task<MVPDetail> VisitOnePage(string url, Dictionary<string, CategoryDetail> categories)
    {
        var mvpdetail = new MVPDetail();

        if (!url.StartsWith("http"))
        {
            url = "https://mvp.microsoft.com" + url;
        }
        mvpdetail.Url = url;
        await Page.GotoAsync(url);
        
        var content = await Page.Locator(".infoContent").AllTextContentsAsync();
        var category = content[0].Trim();
        var images = await Page.Locator(".photoPanel img").AllAsync();
        var name = (await images[2].GetAttributeAsync("alt")).Replace(" photo", "");

        if (name.IndexOf("(") < 0)
        {
            mvpdetail.Name_En = name.Substring(0);
            mvpdetail.Name_Cn = null;
        }
        else
        {
            mvpdetail.Name_En = name.Substring(0, name.IndexOf("("));
            mvpdetail.Name_Cn = name.Substring(name.IndexOf("(") + 1, name.Length- name.IndexOf("(") - 2);
        }

        mvpdetail.PhotoUrl = "https://mvp.microsoft.com"+(await images[2].GetAttributeAsync("src"));

        if (await Page.Locator(".state").IsVisibleAsync())
        {
            mvpdetail.City = await Page.Locator(".state").TextContentAsync();
        }
        var infoContent = await Page.Locator(".infoContent").AllTextContentsAsync();

        mvpdetail.Category = infoContent[0];
        mvpdetail.SinceYear = infoContent[1];
        mvpdetail.NumberOfMVP = infoContent[2];
        var extraInfo = await Page.Locator(".otherRow").AllAsync();
        if (extraInfo.Count > 0)
        {
            var innerHtml=(await extraInfo[0].Locator(".otherContent").InnerHTMLAsync());
            innerHtml=Regex.Replace(innerHtml, @"<[^>]*>", string.Empty);
            if (innerHtml.Contains("Email"))
            {
                mvpdetail.Email =  innerHtml.Replace("Email:", "").Replace("\n", "").Replace(" ", "").Trim();
            }
        }
        if (await Page.Locator(".biography .content").IsVisibleAsync())
        {
            var bio = await Page.Locator(".biography .content").TextContentAsync();
            mvpdetail.Biography = bio.Replace("\n", "").Trim();
        }
        if (categories.ContainsKey(category))
        {
            var detail=categories[category];
            detail.Increase();
            detail.Names.Add(name);
        }
        else
        {
            categories.Add(category, new CategoryDetail() { Category= category, Count=1 });
        }
        return mvpdetail;
    }
}
