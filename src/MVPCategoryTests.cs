using Ganss.Excel;
using Microsoft.Playwright.NUnit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MVPSpider
{
    public class MVPCategoryTests : PageTest
    {
        private async Task GetUrlsFromOneUrl(string pageUrl, List<string> UrlList)
        {
            await Page.GotoAsync(pageUrl);
            var links = await Page.Locator(".profileListItem a").AllAsync();
            foreach (var link in links)
            {
                var url = await link.GetAttributeAsync("href");
                if (!UrlList.Contains(url))
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
                mvpdetail.Name_Cn = name.Substring(name.IndexOf("(") + 1, name.Length - name.IndexOf("(") - 2);
            }

            mvpdetail.PhotoUrl = "https://mvp.microsoft.com" + (await images[2].GetAttributeAsync("src"));

            /*if (await Page.Locator(".state").IsVisibleAsync())
            {
                mvpdetail.City = await Page.Locator(".state").TextContentAsync();
            }*/
            var infoContent = await Page.Locator(".infoContent").AllTextContentsAsync();

            mvpdetail.Category = infoContent[0];
            //mvpdetail.SinceYear = infoContent[1];
            mvpdetail.YearInProgram = infoContent[2];
            /*
            var extraInfo = await Page.Locator(".otherRow").AllAsync();
            if (extraInfo.Count > 0)
            {
                var innerHtml = (await extraInfo[0].Locator(".otherContent").InnerHTMLAsync());
                innerHtml = Regex.Replace(innerHtml, @"<[^>]*>", string.Empty);
                if (innerHtml.Contains("Email"))
                {
                    mvpdetail.Email = innerHtml.Replace("Email:", "").Replace("\n", "").Replace(" ", "").Trim();
                }
            }*/
            if (await Page.Locator(".biography .content").IsVisibleAsync())
            {
                var bio = await Page.Locator(".biography .content").TextContentAsync();
                mvpdetail.Biography = bio.Replace("\n", "").Trim();
            }
            if (categories.ContainsKey(category))
            {
                var detail = categories[category];
                detail.Increase();
                detail.Names.Add(name);
            }
            else
            {
                categories.Add(category, new CategoryDetail() { Category = category, Count = 1 });
            }
            return mvpdetail;
        }
    }
}
