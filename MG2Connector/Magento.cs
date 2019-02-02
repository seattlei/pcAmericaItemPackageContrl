using System;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using System.Drawing;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing.Imaging;
using System.Data.SqlClient;

namespace MG2Connector
{
    public class Magento
    {
        private RestClient Client { get; set; }
        private string Token { get; set; }
        public Int32 errorCount = 0;
        
        public Magento(string magentoUrl)
        {
            Client = new RestClient(magentoUrl);           
        }
        public Magento(string magentoUrl,string token)
        {
            Client = new RestClient(magentoUrl);
            Token = token;
        }

        public string GetAdminToken(string userName, string passWord)
        {
            var request = CreateRequest("/rest/V1/integration/admin/token", Method.POST);
            var user = new Credentials();
            user.username = userName;
            user.password = passWord;

            string json = JsonConvert.SerializeObject(user, Formatting.Indented);

            request.AddParameter("application/json", json, ParameterType.RequestBody);

            var response = Client.Execute(request);
            if(response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return response.Content;
            }else
            {
                return "";
            }
        }

        private RestRequest CreateRequest(string endPoint,Method method)
        {
            var request = new RestRequest(endPoint, method);
            request.RequestFormat = DataFormat.Json;
            return request;
        }

        public void GetSku(string token,string sku)
        {
            var request = CreateRequest("/rest/V1/catalogCategoryListV1GetListGet", Method.GET, token);

            var response = Client.Execute(request);

            if(response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                M2ProductGET product = JsonConvert.DeserializeObject<M2ProductGET>(response.Content);

            }

        }

        public M2ProductGET GetProductlist(string token)
        {
            var request = CreateRequest("/rest/default/V1/products/Z ICE PODS  1", Method.GET, token);

            var response = Client.Execute(request);

            M2ProductGET product = null;
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                product = JsonConvert.DeserializeObject<M2ProductGET>(response.Content);
                return product;
            }
            else {
                return product;
            }

        }

        public M2CategoryGET GetCategoryList(string token)
        {
            var request = CreateRequest("/rest/V1/categories", Method.GET, token);

            var response = Client.Execute(request);
            
            M2CategoryGET Category = null ;
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                 Category = JsonConvert.DeserializeObject<M2CategoryGET>(response.Content);
               

                return Category;
            }
            else {
                //return JsonConvert.DeserializeObject<M2CategoryGET>(response.Content);
                return null;
            }
            
        }

        public void CreateCategory(string categoryName)
        {
            var request = CreateRequest("/rest/V1/categories", Method.POST, Token);
            var cat = new ProductCategory();
            var category = new M2CategoryPOST();
            category.Name = categoryName;

            cat.Category = category;

            string json = JsonConvert.SerializeObject(cat, Formatting.Indented);

            request.AddParameter("application/json", json, ParameterType.RequestBody);

            var response = Client.Execute(request);
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                //return response.Content;
            }
            else
            {
                //return "";
            }

        }


        public string CreateProduct(string productSKU, M2CreateProducts product )
        {
            var request = CreateRequest("/rest/V1/products/"+ productSKU, Method.PUT, Token);
            //var request = CreateRequest("/rest/V1/products/" + productSKU, Method.GET, Token);
            //var product = new M2CreateProducts();
            Stopwatch stopWatch = new Stopwatch();


            string json = JsonConvert.SerializeObject(product, Formatting.Indented);

            request.AddParameter("application/json", json, ParameterType.RequestBody);
            stopWatch.Start();
            var response = Client.Execute(request);
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",ts.Hours, ts.Minutes, ts.Seconds,ts.Milliseconds / 10);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                elapsedTime = elapsedTime + " - UPDATED";
                return "OK";
            }
            else
            {
                String aaa = response.Content;
                if (response.Content == "{\"message\":\"URL key for specified store already exists.\"}")
                {
                    elapsedTime = elapsedTime + " - URL key already exists.";
                    return "Error";
                }
                else if(response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {//Unauthorized
                    elapsedTime = elapsedTime + " - Error";
                    return "TimeOut";
                }
                    
            }
            //return response.Content;
            //return elapsedTime;
            return "Error";

        }
        public string UpdateOneProduct(string productSKU, string connectionString)
        {
            var product = new M2CreateProducts();
            var proDucts = new Product();
            var productQTY = new StockItem();
            var productExtension = new ExtensionAttributes();
            List<CustomAttribute> NewCustomAttribute = new List<CustomAttribute>();
            var custom_attribute_url_key = new CustomAttribute();

            var request = CreateRequest("/rest/V1/products/" + productSKU, Method.PUT, Token);
            List<ProductLink> newProductLink = new List<ProductLink>();

            string strSQLString = " SELECT   [ItemNum],[ItemName_Extra],Price,In_Stock" +
                                " FROM [Inventory] where [ItemNum]= '" + productSKU + "'";

            SqlConnection adoCon = new SqlConnection();
            adoCon.ConnectionString = connectionString;
            adoCon.Open();
            SqlCommand command = new SqlCommand(strSQLString, adoCon);
            SqlDataReader reader = command.ExecuteReader();
            

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    proDucts.Sku = productSKU;
                    proDucts.Name = reader["ItemName_Extra"].ToString().Trim();    // ItemName_Extra
                    proDucts.AttributeSetId = 4;
                    proDucts.Price = Convert.ToSingle(reader["Price"].ToString().Trim());
                    proDucts.Status = 1;
                    proDucts.Visibility = 1;
                    proDucts.TypeId = "simple";

                    custom_attribute_url_key.AttributeCode = "url_key";
                    custom_attribute_url_key.Value = productSKU;
                    NewCustomAttribute.Add(custom_attribute_url_key);
                    proDucts.CustomAttributes = NewCustomAttribute;

                    product.Product = proDucts;
                }
            }

            reader.Close();

            string json = JsonConvert.SerializeObject(product, Formatting.Indented);

            request.AddParameter("application/json", json, ParameterType.RequestBody);
            var response = Client.Execute(request);
            
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return "OK";
            }
            else
            {
                return "TimeOut";
            }
           

        }
        public string Read_a_Product(string productSKU)
        {
            //Stopwatch stopWatch = new Stopwatch();

            var request = CreateRequest("/rest/V1/products/" + productSKU , Method.GET, Token);
            //stopWatch.Start();
            var response = Client.Execute(request);
            //stopWatch.Stop();
            //TimeSpan ts = stopWatch.Elapsed;
            //string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);

           
             return response.Request.ToString();
           
        }

        public RestRequest CreateRequest(string endPoint, Method method,string token)
        {
            var request = new RestRequest(endPoint, method);
            request.RequestFormat = DataFormat.Json;
            request.AddHeader("Authorization", "Bearer " + token);
            request.AddHeader("Accept", "application/json");
            return request;
        }

        public string ConvertImage(string fileName)
        {
            Bitmap image = null;
            FileStream fs = null;

            try
            {

                if (File.Exists(@"C:\hibroweb\" + fileName + ".jpg"))
                {
                    fs = new FileStream(@"C:\hibroweb\" + @fileName + ".jpg", FileMode.Open);
                    image = (Bitmap)Bitmap.FromStream(fs);
                    //fs.Close();

                }
                else if (File.Exists(@"C:\hibroweb\" + fileName + ".png"))
                {
                    fs = new FileStream(@"C:\hibroweb\" + @fileName + ".png", FileMode.Open);
                    image = (Bitmap)Bitmap.FromStream(fs);
                    //fs.Close();
                }
                else
                {
                    return "No File";
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    // Convert Image to byte[]
                    image.Save(ms, ImageFormat.Png);
                    byte[] imageBytes = ms.ToArray();

                    // Convert byte[] to Base64 String
                    string base64String = Convert.ToBase64String(imageBytes);
                    fs.Close();
                    return base64String;
                }

            }
            catch (Exception ex)
            {
                return "Error";
            }
        }

        public void MoveImageFile(string fileName, string departmentID)
        {
            try
            {
                string targetPath =@"C:\hibroweb\" + departmentID.Replace("/","-");
                if (!System.IO.Directory.Exists(targetPath))
                {
                    System.IO.Directory.CreateDirectory(targetPath);
                }
                string sourceFile = @"C:\hibroweb\" + fileName + ".jpg";
                string destinationFile = @targetPath + "\\" + fileName + ".jpg";

                if (File.Exists(sourceFile))
                {
                    System.IO.File.Move(sourceFile, destinationFile);

                }
                sourceFile = @"C:\\hibroweb\\" + fileName + ".png";
                destinationFile = @targetPath + "\\" + fileName + ".png";
                if (File.Exists("C:\\hibroweb\\" + fileName + ".png"))
                {
                    System.IO.File.Move(sourceFile, destinationFile);

                }

            }
            catch (Exception ex)
            {
                //return "Error";
            }


        }
    }
}
