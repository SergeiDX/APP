using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APP
{
    internal class conexion
    {
        public SqlConnection cnn = new SqlConnection(@"Server=DESKTOP-EUB75BK;Database=Capacitacion;Integrated Security=True");
    }
}
