using Ganss.Excel;
using Microsoft.Playwright;
using Microsoft.Playwright.NUnit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

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

            var count=((JsonElement)jsonEle).GetProperty("filteredCount").GetInt32();
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
        [TearDown]
        public async Task TearDownAPITesting()
        {
            await APIRequest.DisposeAsync();
        }
    }
}
