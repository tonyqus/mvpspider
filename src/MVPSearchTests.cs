using Ganss.Excel;
using Microsoft.Playwright;
using Microsoft.Playwright.NUnit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace MVPSpider
{
    public class MVPSearchBody
    {
        public MVPSearchBody(string program, string country, int pageSize = 20) {
            this.program = new();
            this.program.Add(program);
            this.countryRegionList = new();
            if (country != null)
            {
                this.countryRegionList.Add(country);
            }
            this.stateProvinceList = new();
            this.technicalExpertiseList = new();
            this.technologyFocusAreaGroupList = new();
            this.technologyFocusAreaList = new();
            this.milestonesList = new();
            this.languagesList = new();
            this.industryFocusList = new();
            this.pageSize = pageSize;
            this.pageIndex = 1;
        }
        public string searchKey { get; set; }
        public string academicInstitution { get; set; }
        public List<string> program { get; set; }
        public List<string> countryRegionList { get; set; }
        public List<string> stateProvinceList { get; set; }
        public List<string> technicalExpertiseList { get; set; }
        public List<string> technologyFocusAreaGroupList { get; set; }
        public List<string> technologyFocusAreaList { get; set; }
        public List<string> milestonesList { get; set; }
        public List<string> languagesList { get; set; }
        public List<string> industryFocusList { get; set; }
        public int pageSize { get; set; }
        public int pageIndex { get; set; }

    }
    public class MVPCountryCount
    { 
        public string Name { get; set; }
        public int Count { get; set; }
    }
    public class MVPSearchTests:PlaywrightTest
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
        public async Task GetCount_MVP_China()
        {
            var body = new MVPSearchBody("MVP", "China");
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject=body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count=jsonEle?.GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(144));
        }
        [Test]
        public async Task GetCount_MSP_China()
        {
            var body = new MVPSearchBody("MSP", "China");
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count = ((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(87));
        }
        [Test]
        public async Task GetCount_RD_China()
        {
            var body = new MVPSearchBody("RD", "China");
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count = ((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(9));
        }
        [Test]
        public async Task GetCount_MVP_Taiwan()
        {
            var body = new MVPSearchBody("MVP", "Taiwan");
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count = ((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(49));
        }
        [Test]
        public async Task GetCount_MVP_Global()
        {
            var body = new MVPSearchBody("MVP", null);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count = ((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(3179));
        }
        [Test]
        public async Task GetCount_MSP_Global()
        {
            var body = new MVPSearchBody("MSP", null);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            var count = ((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
            Assert.That(count, Is.EqualTo(3591));
        }
        private async Task GetCountryCount(string program, string filename)
        {
            var request = await this.FrontRequest.GetAsync("/Locales/CountryOrRegion.json");
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();
            var names = ((JsonElement)jsonEle).EnumerateObject()
                .Select(o => o.Value.ToString()).ToList();

            Assert.That(names.Count, Is.EqualTo(247));
            List<MVPCountryCount> countryCounts = new List<MVPCountryCount>();
            var i = 1;
            foreach (var name in names)
            {
                var body = new MVPSearchBody(program, name);
                var request2 = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
                var jsonEle2 = await request2.JsonAsync();
                var count = ((JsonElement)jsonEle2).GetProperty("filteredCount").GetInt32();
                if (i % 10 == 0)
                {
                    TestContext.Progress.WriteLine($"{i} records done");
                }
                if (count > 0)
                    countryCounts.Add(new MVPCountryCount() { Name = name, Count = count });
                i++;
            }
            new ExcelMapper().Save(filename, countryCounts, "Country Counts");
        }
        [Test]
        public async Task GetMVP_Counts()
        {
            await GetCountryCount("MVP", "mvp_country_counts.xlsx");
        }
        [Test]
        public async Task GetMSP_Counts()
        {
            await GetCountryCount("MSP", "msp_country_counts.xlsx");
        }
        [Test]
        public async Task GetRegionalDirector_Counts()
        {
            await GetCountryCount("RD", "rd_country_counts.xlsx");
        }
        [Test]
        public async Task GetChinaMVPUrls()
        {
            var body = new MVPSearchBody("MVP", "China",500);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var jsonEle = await request.JsonAsync();

            //var count = jsonEle?.GetProperty("filteredCount").GetInt32();
            var mvpIds = jsonEle?.Get("communityLeaderProfiles")?.EnumerateArray().Select(x=>x.Get("userProfileIdentifier")?.GetString()).ToList();

            Assert.That(mvpIds?.Count, Is.EqualTo(145));
        }
        const string mvpSingleApiUrl = "/api/mvp/UserProfiles/public/";
        const string rdSingleApiUurl = "/api/rd/UserProfiles/public/";
        [Test]
        public async Task GetMVPDetailListAndSaveToExcel_Taiwan()
        {
            var body = new MVPSearchBody("MVP", "Taiwan", 200);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var mvpDetails = new List<MVPDetail>();

            var jsonEle = await request.JsonAsync();
            var UrlList = jsonEle?.Get("communityLeaderProfiles")?.EnumerateArray().ToDictionary(
                  x => x.Get("userProfileIdentifier")?.GetString() + string.Empty,
                  x => mvpSingleApiUrl + x.Get("userProfileIdentifier")?.GetString()
                );

            foreach (var url in UrlList)
            {
                var MVPDetail = await VisitOnePage(url.Key, url.Value);
                mvpDetails.Add(MVPDetail);
            }
            new ExcelMapper().Save("mvp_taiwan.xlsx", mvpDetails, "MVPs");
        }
        [Test]
        public async Task GetMVPDetailListAndSaveToExcel_ChinaMainland()
        {
            var body = new MVPSearchBody("MVP", "China", 500);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var mvpDetails = new List<MVPDetail>();

            var jsonEle = await request.JsonAsync();
            var UrlList = jsonEle?.Get("communityLeaderProfiles")?.EnumerateArray().ToDictionary(
                  x=> x.Get("userProfileIdentifier")?.GetString()+string.Empty,
                  x=>mvpSingleApiUrl + x.Get("userProfileIdentifier")?.GetString()
                );
            foreach (var url in UrlList)
            {
                var MVPDetail = await VisitOnePage(url.Key, url.Value);
                mvpDetails.Add(MVPDetail);
            }
            new ExcelMapper().Save("mvp_china.xlsx", mvpDetails, "MVPs");
        }
        [Test]
        public async Task GetRDDetailListAndSaveToExcel_Global()
        {
            var body = new MVPSearchBody("RD", null, 300);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var mvpDetails = new List<MVPDetail>();

            var jsonEle = await request.JsonAsync();
            var UrlList = jsonEle?.Get("communityLeaderProfiles")?.EnumerateArray().ToDictionary(
                  x => x.Get("userProfileIdentifier")?.GetString() + string.Empty,
                  x => rdSingleApiUurl + x.Get("userProfileIdentifier")?.GetString()
                );
            foreach (var url in UrlList)
            {
                var MVPDetail = await VisitOnePage(url.Key, url.Value, true);
                if (MVPDetail != null)
                    mvpDetails.Add(MVPDetail);
            }
            new ExcelMapper().Save("rd_global.xlsx", mvpDetails, "MVPs");
        }
        [Test]
        public async Task GetMVPDetailListAndSaveToExcel_Global()
        {
            var body = new MVPSearchBody("MVP", null, 4000);
            var request = await this.APIRequest.PostAsync("/api/CommunityLeaders/search/", new() { DataObject = body });
            Assert.True(request.Ok);
            var mvpDetails = new List<MVPDetail>();

            var jsonEle = await request.JsonAsync();
            var UrlList = jsonEle?.Get("communityLeaderProfiles")?.EnumerateArray().ToDictionary(
                  x => x.Get("userProfileIdentifier")?.GetString() + string.Empty,
                  x => mvpSingleApiUrl + x.Get("userProfileIdentifier")?.GetString()
                );

            foreach (var url in UrlList)
            {
                var MVPDetail = await VisitOnePage(url.Key, url.Value);
                if(MVPDetail!=null)
                    mvpDetails.Add(MVPDetail);
            }
            new ExcelMapper().Save("mvp_global.xlsx", mvpDetails, "MVPs");
        }
        private async Task<MVPDetail> VisitOnePage(string mvpguid, string apiUrl, bool isRD = false)
        {
            var mvpdetail = new MVPDetail();
            if (!isRD)
            {
                mvpdetail.Url = "https://mvp.microsoft.com/en-US/mvp/profile/" + mvpguid;
            }
            else
            {
                mvpdetail.Url = "https://mvp.microsoft.com/en-US/RD/profile/" + mvpguid;
            }
            try
            {
                var request = await this.APIRequest.GetAsync(apiUrl);
                if (!request.Ok)
                {
                    return null;
                }

                var jsonEle = await request.JsonAsync();
                var node = jsonEle?.Get("userProfile");
                mvpdetail.Country = node?.Get("addressCountryOrRegionName")?.GetString();
                mvpdetail.Name_En = node?.Get("firstName")?.GetString() + " " + node?.Get("lastName")?.GetString();
                if (node?.Get("localizedFirstName")?.GetString() != null)
                {
                    mvpdetail.Name_Cn = node?.Get("localizedFirstName")?.GetString() + " " + node?.Get("localizedLastName")?.GetString();
                }

                mvpdetail.PhotoUrl = node?.Get("profilePictureUrl")?.GetString();

                if (!isRD)
                {
                    mvpdetail.Category = string.Join(",", node?.Get("awardCategory")?.EnumerateArray().Select(x => x.GetString()).ToArray());
                }
                if (!isRD)
                {
                    mvpdetail.TechFocus = string.Join(",", node?.Get("technologyFocusArea")?.EnumerateArray().Select(x => x.GetString()).ToArray());
                }
                else
                {
                    mvpdetail.TechFocus = string.Join(",", node?.Get("technicalExpertise")?.EnumerateArray().Select(x => x.GetString()).ToArray());
                }
                mvpdetail.YearInProgram = node?.Get("yearsInProgram")?.GetInt32().ToString();

                var titleName = node?.Get("titleName")?.GetString();
                if (titleName != null)
                {
                    if (titleName.StartsWith("Ms") || titleName.StartsWith("Mrs"))
                    {
                        mvpdetail.Gender = "F";
                    }
                    else
                    {
                        mvpdetail.Gender = "M";
                    }
                }
                mvpdetail.CompanyName = node?.Get("companyName")?.GetString();
                mvpdetail.Biography = node?.Get("biography")?.GetString()?.Replace("\n", "").Trim();
                var links = node?.Get("userProfileSocialNetwork")?.EnumerateArray().ToList();
                foreach (var link in links)
                {
                    var socialNetworkName = link.Get("socialNetworkName")?.GetString()?.ToLower();
                    switch (socialNetworkName)
                    {
                        case "github":
                            mvpdetail.Social_Github = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "linkedin":
                            mvpdetail.Social_Linkedin = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "facebook":
                            mvpdetail.Social_Facebook = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "twitter":
                            mvpdetail.Social_Twitter = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "youtube":
                            mvpdetail.Social_Youtube = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "personal website":
                            mvpdetail.Social_Blog = link.Get("socialNetworkImageLink")?.GetString();
                            break;
                        case "other":
                            var url = link.Get("socialNetworkImageLink")?.GetString();
                            if (url.Contains("bilibili.com"))
                            {
                                mvpdetail.Social_Bilibili = url;
                            }
                            else if (url.Contains("cnblogs.com") && mvpdetail.Social_Blog == null)
                            {
                                mvpdetail.Social_Blog = url;
                            }
                            break;
                    }
                }
            }
            catch (PlaywrightException)
            {
                return null;
            }
            return mvpdetail;

        }
        [TearDown]
        public async Task TearDownAPITesting()
        {
            await APIRequest.DisposeAsync();
        }
    }
}
