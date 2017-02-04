using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public enum ExecResult : int
    {
        Failure = 0,
        Success,
        Creating,
        ReLocate, //二次入职
        ReCreate,
        Existing,
        FirstNameEmpty,
        LastNameEmpty,
        DepartmentEmpty
    }
}
