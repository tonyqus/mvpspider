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
        var content = await Page.Locator(".infoContent").AllTextContentsAsync();
        Assert.That(content[0], Is.EqualTo("Developer Technologies"));
        Assert.That(content[1], Is.EqualTo("2021"));
        Assert.That(content[2], Is.EqualTo("3"));

        //get contact and language info
        var extraInfo = await Page.Locator(".otherContent").AllTextContentsAsync();
        Assert.That(extraInfo[0].Replace("\n", "").Trim(), Is.EqualTo("Email:                            liuliang79@live.com"));
        Assert.That(extraInfo[1].Replace("\n", "").Trim(), Is.EqualTo("Chinese - Simplified"));

        //get location
        var country = await Page.Locator(".country").TextContentAsync();
        Assert.That(country, Is.EqualTo("China"));
        var state = await Page.Locator(".state").TextContentAsync();
        Assert.That(state, Is.EqualTo("Beijing"));

        //get biography
        var bio = await Page.Locator(".biography .content").TextContentAsync();
        Assert.IsTrue(bio.Replace("\n", "").Trim().Length > 0);

        //recent activities
        var activities = await Page.Locator(".raListTable tr").AllAsync();
        Assert.IsTrue(activities.Count > 0);
    }
    [Test]
    public async Task GetMVPSinglePageContent_XinyueTang()
    {
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/PublicProfile/5001884?fullName=Xinyue%20Tang");
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

        var images = await Page.Locator(".photoPanel img").AllAsync();
        var imgUrl = await images[2].GetAttributeAsync("src");
        var name = (await images[2].GetAttributeAsync("alt")).Replace(" photo","");
        Assert.IsNotEmpty(imgUrl);
        Assert.That(name, Is.EqualTo("James Yeung (杨舜杰)"));
        //get mvp category and year
        var content = await Page.Locator(".infoContent").AllTextContentsAsync();
        Assert.That(content[0], Is.EqualTo("Developer Technologies"));
        Assert.That(content[1], Is.EqualTo("2020"));
        Assert.That(content[2], Is.EqualTo("3"));

        //get contact and language info
        var extraInfo = await Page.Locator(".otherRow").AllAsync();
        var otherInfo = new Dictionary<string, string>();
        foreach (var info in extraInfo)
        {
            var category = await info.Locator(".otherCatalog").TextContentAsync();
            var detail = await info.Locator(".otherContent").InnerHTMLAsync();
            otherInfo.Add(category.Trim(), detail.Trim());
        }
        Assert.That(otherInfo.Count, Is.EqualTo(3));
        //get location
        var country = await Page.Locator(".country").TextContentAsync();
        Assert.That(country, Is.EqualTo("China"));
        var state = await Page.Locator(".state").TextContentAsync();
        Assert.That(state, Is.EqualTo("Shanghai"));

        //get biography
        var bio = await Page.Locator(".biography .content").TextContentAsync();
        Assert.IsTrue(bio.Replace("\n", "").Trim().Length > 0);

        //recent activities
        var activities = await Page.Locator(".raListTable tr").AllAsync();
        Assert.IsTrue(activities.Count > 0);
    }
    [Test]
    public async Task GetMVPCount_China()
    {
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/MvpSearch?lo=China&sc=e");
        var total = await Page.Locator(".resultcount").TextContentAsync();
        var numStr = total.Replace("(", "").Replace(")", "");
        Assert.That(Int32.Parse(numStr), Is.EqualTo(138));
    }
    [Test]
    public async Task GetMVPCount_Taiwan()
    {
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/MvpSearch?lo=Taiwan&sc=e");
        var total = await Page.Locator(".resultcount").TextContentAsync();
        var numStr = total.Replace("(", "").Replace(")", "");
        Assert.That(Int32.Parse(numStr),Is.EqualTo(40));
    }
    [Test]
    public async Task GetMVPCount_Global()
    {
        //global mvp count
        await Page.GotoAsync("https://mvp.microsoft.com/en-us/MvpSearch?kw=&x=2&y=11");
        var total = await Page.Locator(".resultcount").TextContentAsync();
        var numStr = total.Replace("(", "").Replace(")", "");
        Assert.That(Int32.Parse(numStr),Is.EqualTo(3000));
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
