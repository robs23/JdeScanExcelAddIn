using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class Place : Entity<Place>
    {
        [DisplayName("ID")]
        public int PlaceId { get; set; }
        public override int Id
        {
            set => value = PlaceId;
            get => PlaceId;
        }
        public string Name { get; set; }

        public bool? IsArchived { get; set; }
    }
}
