using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Import.Entities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using MySqlConnector;
using Dapper;
using System.Data;

namespace MISA.Import.Controllers
{
    [Route("api/v1/customers")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        // Khởi tạo kết nối đến data base
        IDbConnection dbConnection;
        // Thông tin data base
        string connectionString = ""
                + "Host=47.241.69.179;"
                + "Port=3306;"
                + "User Id=dev;"
                + "Password=12345678;"
                + "Database = MF827_Import_PBAC;";

        // Lấy dữ liệu từ file excel từ phía người dùng và trả lại dữ liệu dạng JSON
        // <param name="formFile">File dữ liệu</param>
        // <returns>
        // 200 - lấy dữ liệu thành công
        // 400 - dữ liệu đầu vào không hợp lệ (file trống hoặc không có file đầu vào)
        // 500 - có lỗi xảy ra phía server (exception,...)
        // </returns>
        [HttpPost("import")]
        public IActionResult Import(IFormFile formFile)
        {
            // Kiểm tra file có hợp lệ hay không
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest();
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest();
            }

            // Nếu có thì lấy toàn bộ bản ghi từ file
            var customers = GetAllRecordFromFile(formFile);
            // Validate toàn bộ bản ghi
            Validate(customers);
 
            // Trả về số lượng bản ghi hợp lệ, không hợp lệ và khách hàng
            var result = new
            {
                numberOfValidation = CountNumberValidation(customers),
                numberOfInvalidate = customers.Count - CountNumberValidation(customers),
                data = customers
            };

            return Ok(result);
        }

        // Thêm mới list khách hàng
        // <param name="customers">Thông tin đối tượng</param>
        // <returns>
        // 201 - thêm mới thành công
        // 204 - không thêm được vào database
        // 400 - dữ liệu đầu vào không hợp lệ
        // 500 - có lỗi xảy ra phía server (exception,...)
        // </returns>
        [HttpPost]
        public IActionResult Post(List<Customer> customers)
        {
            var res = AddToDatabase(customers);
            if (res > 0)
            {
                return StatusCode(201, res);
            }
            else
            {
                return NoContent();
            }
        }

        // Lấy toàn bộ bản ghi từ file và trả về dữ liệu khách hàng dưới dạng List
        private List<Customer> GetAllRecordFromFile(IFormFile formFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var customers = new List<Customer>();

            using (var stream = new MemoryStream())
            {
                formFile.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    for (int row = 3; row < worksheet.Dimension.Rows; row++)
                    {
                        string dobString = worksheet.Cells[row, 6].Value == null ? null : worksheet.Cells[row, 6].Value.ToString();
                        Customer c = new Customer(
                            // CustomerId
                            new Guid(),
                            // CustomerCode
                            worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString(),
                            // FullName
                            worksheet.Cells[row, 2].Value == null ? null : worksheet.Cells[row, 2].Value.ToString(),
                            // MemberCardCode
                            worksheet.Cells[row, 3].Value == null ? null : worksheet.Cells[row, 3].Value.ToString(),
                            // CustomerGroupId
                            null,
                            // CustomerGroupName
                            worksheet.Cells[row, 4].Value == null ? null : worksheet.Cells[row, 4].Value.ToString(),
                            // PhoneNumber
                            worksheet.Cells[row, 5].Value == null ? null : worksheet.Cells[row, 5].Value.ToString(),
                            // Date of birth
                            ParseDate(dobString),
                            // CompanyName
                            worksheet.Cells[row, 7].Value == null ? null : worksheet.Cells[row, 7].Value.ToString(),
                            // CompanyTaxCode
                            worksheet.Cells[row, 8].Value == null ? null : worksheet.Cells[row, 8].Value.ToString(),
                            // Email
                            worksheet.Cells[row, 9].Value == null ? null : worksheet.Cells[row, 9].Value.ToString(),
                            // Address
                            worksheet.Cells[row, 10].Value == null ? null : worksheet.Cells[row, 10].Value.ToString(),
                            // Note
                            worksheet.Cells[row, 11].Value == null ? null : worksheet.Cells[row, 11].Value.ToString());
                        customers.Add(c);
                    }
                }
            }
            return customers;
        }

        // Xử lý ngày sinh
        DateTime? ParseDate(string dob)
        {
            try
            {
                var elements = dob.Split("/");
                // Nếu chỉ có năm thì tự động điền 01/01/[năm]
                if (elements.Length == 1)
                {
                    return new DateTime(int.Parse(elements[0]), 1, 1);
                }
                // Nếu chỉ có tháng và năm thì tự động điền 01/[tháng]/[năm]
                else if (elements.Length == 2)
                {
                    return new DateTime(int.Parse(elements[1]), int.Parse(elements[0]), 1);
                }
                // Đầy đủ ngày tháng năm
                else
                {
                    return new DateTime(int.Parse(elements[2]), int.Parse(elements[1]), int.Parse(elements[0]));
                }
            }
            catch (Exception e)
            {
                return null;
            }
        }

        // Hàm thêm khách hàng dưới dạng list vào data base
        // <return>Số lượng bản ghi lưu vào data base thành công</return>
        private int AddToDatabase(List<Customer> customers)
        {
            // Khởi tạo tham số số lượng bản ghi
            int rowsAffect = 0;
            foreach (Customer customer in customers)
            {
                // Kiểm tra status của khách hàng có trống hoặc rỗng không
                if(string.IsNullOrEmpty(customer.Status))
                {
                    // Nếu có thì thêm vào data base
                    using (dbConnection = new MySqlConnection(connectionString))
                    {
                        // Tăng số lượng bản ghi
                        rowsAffect += dbConnection.Execute("Proc_InsertCustomer", param: customer, commandType: CommandType.StoredProcedure);
                    }
                }
            }
            return rowsAffect;
        }

        // Validate dữ liệu
        private void Validate(List<Customer> customers)
        {
            foreach(Customer customer in customers)
            {
                // Kiểm tra dữ liệu trùng lặp của đối tượng khách hàng với dữ liệu trong data base
                CheckDuplicateInDatabase(customer);
                // Kiểm tra nhóm khách hàng có tồn tại không
                CheckCustomerGroupExist(customer);
            }


            // Kiểm tra mã khách hàng trong file
            for(int i = 0; i < customers.Count; i++)
            {
                // Kiểm tra mã khách hàng có trống hoặc null không
                if (string.IsNullOrEmpty(customers[i].CustomerCode)) {
                    customers[i].Status += "Mã khách hàng không được phép để trống\n";
                }
                else
                {
                    // Kiểm tra mã khách hàng trùng lặp trong file
                    for (int j = 0; j < customers.Count; j++)
                    {
                        if (i != j)
                        {
                            if (customers[i].CustomerCode == customers[j].CustomerCode)
                            {
                                customers[i].Status += "Mã khách hàng trùng lặp trên file\n";
                                break;
                            }
                        }
                    }
                }
            }

            // Kiểm tra số điện thoại trong file
            for (int i = 0; i < customers.Count; i++)
            {
                // Kiểm tra số điện thoại có trống hoặc null không
                if (string.IsNullOrEmpty(customers[i].PhoneNumber)) {
                    customers[i].Status += "Số điện thoại không được phép để trống\n";
                }
                else
                {
                    // Kiểm tra số điện thoại trùng lặp trong file
                    for (int j = 0; j < customers.Count; j++)
                    {
                        if (i != j)
                        {
                            if (customers[i].PhoneNumber == customers[j].PhoneNumber)
                            {
                                customers[i].Status += "Số điện thoại trùng lặp trên file\n";
                                break;
                            }
                        }
                    }
                }
                
            }

            // Kiểm tra email trùng lặp trong file
            for (int i = 0; i < customers.Count; i++)
            {
                // Kiểm tra email có trống hoặc null không
                if (string.IsNullOrEmpty(customers[i].Email)) {
                    customers[i].Status += "Email không được phép để trống\n";
                }
                else
                {
                    // Kiểm tra email có trùng lặp trong file
                    for (int j = 0; j < customers.Count; j++)
                    {
                        if (i != j)
                        {
                            if (customers[i].Email == customers[j].Email)
                            {
                                customers[i].Status += "Email trùng lặp trên file\n";
                                break;
                            }
                        }
                    }
                }
                
            }
        }

        // Kiểm tra nhóm khách hàng có tồn tại không
        private void CheckCustomerGroupExist(Customer customer)
        {
            string sqlCommand = "Proc_CheckCustomerGroupExist";

            DynamicParameters parameters = new DynamicParameters();

            parameters.Add("@Name", customer.CustomerGroupName);


            using(dbConnection = new MySqlConnection(connectionString))
            {
                var cg = dbConnection.QueryFirstOrDefault<CustomerGroup>(sqlCommand, param: parameters, commandType: CommandType.StoredProcedure);
                if(cg == null)
                {
                    customer.Status += "Nhóm khách hàng không tồn tại\n";
                }
                else
                {
                    customer.CustomerGroupId = cg.CustomerGroupId;
                }
            }

            
        }

        // Kiểm tra dữ liệu trùng lặp của đối tượng khách hàng với dữ liệu trong data base
        private void CheckDuplicateInDatabase(Customer customer)
        {
            // Mã khách hàng
            if(CheckAttributeExistInDatabase("CustomerCode", customer.CustomerCode))
            {
                customer.Status += "Mã khách hàng đã tồn tại trên database\n";
            }
            // Số điện thoại
            if (CheckAttributeExistInDatabase("PhoneNumber", customer.PhoneNumber))
            {
                customer.Status += "Số điện thoại đã tồn tại trên database\n";
            }
            // Email
            if (CheckAttributeExistInDatabase("Email", customer.Email))
            {
                customer.Status += "Email đã tồn tại trên database\n";
            }
        }

        // Hàm dùng chung kiểm tra dữ liệu đã tồn tại trong data base
        private bool CheckAttributeExistInDatabase(string attributeName, string value)
        {
            string sqlCommand = $"Proc_Check{attributeName}Exist";

            DynamicParameters parameters = new DynamicParameters();

            parameters.Add($"@{attributeName}", value);

            var exist = true;

            using (dbConnection = new MySqlConnection(connectionString))
            {
                exist = dbConnection.QueryFirstOrDefault<bool>(sqlCommand, param: parameters, commandType: CommandType.StoredProcedure);
            }

            return exist;
        }

        // Hàm đếm số bản ghi không hợp lệ
        private int CountNumberValidation(List<Customer> customers)
        {
            int result = 0;
            for (int i = 0; i < customers.Count; i++)
            {
                if (customers[i].Status == null)
                {
                    result++;
                }
            }
            return result;
        }
    }
}
