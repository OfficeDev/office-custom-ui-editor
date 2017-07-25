using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UndoRedo.Model
{
    class UndoRedoStack : Stack
    {
        public override void Push(string value, int index)
        {
            _list.AddToTail(value, index);
            if (_list.Count > 30) _list.RemoveFromHead(out index);
        }
    }
}
