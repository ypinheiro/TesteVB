using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace APITeste.Models
{
    public class Usuario
    {
        public Int32 Id { get; set; }
        public String Nome { get; set; }
        public String Email { get; set; }
        public Boolean Ativo { get; set; }
    }
}