using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UndoRedo.Model
{
    interface IStack
    {
        bool IsEmpty();
        void Push(string value, int index);
        string Pop(out int index);

        int Count
        {
            get;
        }
    }
}
