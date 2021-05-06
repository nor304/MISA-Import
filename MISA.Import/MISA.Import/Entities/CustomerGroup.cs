using System;
namespace MISA.Import.Entities
{
    public class CustomerGroup
    {
        public CustomerGroup()
        {
        }

        public Guid CustomerGroupId { get; set; }
        public string CustomerGroupName { get; set; }
        public string Description { get; set; }
    }
}
