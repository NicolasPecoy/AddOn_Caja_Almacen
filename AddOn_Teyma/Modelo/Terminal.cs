using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddOn_Caja.Modelo
{
    class Terminal
    {
        public string empresaHash { get; set; }
        public string EmpCod { get; set; }
        public string TermCod { get; set; }


        // podemos recibir hash, empresaCod ,TermCod y Operacion para setearlos al objecto _Transaccion
        public Terminal(string empresaH, string empresaCod, string terminalCod)
        {

            // datos de invenzis tambien configurados en el pos. Cambiaria el Termcod si hay mas de uno
            this.empresaHash = empresaH;
            this.EmpCod = empresaCod;
            this.TermCod = terminalCod;
        }
    }
}
