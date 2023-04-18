using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.PeerResolvers;
using System.Text;

namespace AddOn_Caja.Clases
{
    public class CodRespuestaPOSGeocom
    {
        public CodRespuestaPOSGeocom(string codigo, string mensaje, string respuesta, string estado)
        {
            this.codigo = codigo;
            this.mensaje = mensaje;
            this.respuesta = respuesta;
            this.estado = estado;
        }

        public CodRespuestaPOSGeocom()
        {
      
        }



        public string codigo { get; set; }
        public string mensaje { get; set; }
        public string respuesta { get; set; }
        public string estado { get; set; }



    }
}
