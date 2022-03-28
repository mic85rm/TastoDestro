using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TastoDestro.Model;

namespace TastoDestro.Helper
{
    interface IStorico
    {
        Storico GetStorico(int id);
         void SalvaStorico(Storico storico);
    }
}
