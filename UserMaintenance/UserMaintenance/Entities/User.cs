using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserMaintenance.Entities //egyébként elég lenne innen is kitörölni az entitiest, de nem praktikus
{
    public class User //public, mert szeretném majd máshol is használni
    {
        public Guid ID { get; } = Guid.NewGuid(); //egyedi azonosítót hoz létre a felhasználó szerverének adataiból
        //set ágat törlöm is, nem akarom, hogy megváltoztassák


        //privat rész nem kötelező, hogy létezzen

        public string FullName //nem engedélyezem a szerkesztést, törlöm a setet
        {
            get; set;
       
        }

    }
}
