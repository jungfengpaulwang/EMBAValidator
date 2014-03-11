using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Validator
{
    [Flags]
    public enum OutputOptions
    {
        Correct = 1,
        Full = 2,
        Error = 4
    }
}
