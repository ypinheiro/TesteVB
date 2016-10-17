using System.Data.Entity;

namespace APITeste.Models
{
    public class Contexto: DbContext
    {
        public Contexto()
            : base("name=Connection")
        {
        }

        public DbSet<Usuario> Usuarios { get; set; }
    }
}
