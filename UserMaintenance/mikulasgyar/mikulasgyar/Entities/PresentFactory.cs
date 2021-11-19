using mikulasgyar.Abstractions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mikulasgyar.Entities
{
    class PresentFactory : IToyFactory
    {

        public Color Szalag { get; set; }
        public Color Doboz { get; set; }

        public Toy CreateNew()
        {
            return new Present(Szalag, Doboz);
        }
    }
}
