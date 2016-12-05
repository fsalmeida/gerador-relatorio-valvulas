using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjetoRe.Exceptions
{
    public class CustomException : Exception
    {
        public string Mensagem { get; set; }
        public CustomException(string mensagem)
        {
            this.Mensagem = mensagem;
        }
    }
}
