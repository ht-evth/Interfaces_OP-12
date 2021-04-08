using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Catering_OP_6 {
	class RowInTable {

		public string row_num { get; set; }
		public string card_number { get; set; }

		public string name { get; set; }
		public string code { get; set; }

		public string amount { get; set; }

		public string fact_price { get; set; }
		public string fact_sum { get; set; }

		public string record_price { get; set; }
		public string record_sum { get; set; }

		public string note { get; set; }

		public RowInTable() {
			number = "";
			card_number = "";
			name = "";
			code = "";
			amount = "";
			fact_price = "";
			fact_sum = "";
			record_price = "";
			record_sum = "";
			note = "";
		}
	}
}
