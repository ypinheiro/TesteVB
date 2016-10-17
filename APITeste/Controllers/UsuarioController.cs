using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using APITeste.Models;

namespace APITeste.Controllers
{
    public class UsuarioController : ApiController
    {
        private readonly Contexto _dbContexto = new Contexto();

        // GET api/Usuario
        public HttpResponseMessage Get([FromUri] Usuario item)
        {
            var retorno = _dbContexto.Usuarios.Where(_ => _.Email == item.Email);

            return Request.CreateResponse(HttpStatusCode.OK, retorno);
        }

        // GET api/Usuario/5
        public HttpResponseMessage Get(int value)
        {
            var usuario = _dbContexto.Usuarios.Find(value);

            if (usuario == null)
                return Request.CreateResponse(HttpStatusCode.NotFound);

            return Request.CreateResponse(HttpStatusCode.OK, usuario);
        }

        // PUT api/Usuario/5
        public HttpResponseMessage Put(int id, Usuario usuario)
        {
            if (!ModelState.IsValid)
                return Request.CreateResponse(HttpStatusCode.BadRequest);

            if (id != usuario.Id)
                return Request.CreateResponse(HttpStatusCode.BadRequest);

            try
            {
                _dbContexto.Entry(usuario).State = EntityState.Modified;

                _dbContexto.SaveChanges();
            }
            catch (DbUpdateConcurrencyException ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex);
            }

            return Request.CreateResponse(HttpStatusCode.OK, usuario);
        }

        // POST api/Usuario
        public HttpResponseMessage Post(Usuario usuario)
        {
            try
            {
                if (!ModelState.IsValid)
                    return Request.CreateResponse(HttpStatusCode.BadRequest, ModelState);

                _dbContexto.Usuarios.Add(usuario);
                _dbContexto.SaveChanges();

                var response = Request.CreateResponse(HttpStatusCode.Created, usuario);
                response.Headers.Location = new Uri(Request.RequestUri, string.Format("usuario/{0}", usuario.Id));

                return response;
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex);
            }
        }

        // DELETE api/Usuario/5
        public HttpResponseMessage Delete(int id)
        {
            var usuario = _dbContexto.Usuarios.Find(id);

            if (usuario == null)
                return Request.CreateResponse(HttpStatusCode.NotFound);

            try
            {
                _dbContexto.Usuarios.Remove(usuario);

                _dbContexto.SaveChanges();
            }
            catch (DbUpdateConcurrencyException ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, ex);
            }

            return Request.CreateResponse(HttpStatusCode.OK, usuario);
        }

        protected override void Dispose(bool disposing)
        {
            _dbContexto.Dispose();
            base.Dispose(disposing);
        }
    }
}
