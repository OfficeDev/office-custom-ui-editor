using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UndoRedo.Model
{
    public abstract class Stack : IStack
    {
        protected List _list;

        public Stack()
        {
            _list = new List();
        }

        #region IStack Members

        public bool IsEmpty()
        {
            return _list.IsEmpty();
        }

        public virtual void Push(string value, int index)
        {
            _list.AddToTail(value, index);
        }

        public virtual string Pop(out int index)
        {
            return _list.RemoveFromTail(out index);
        }

        public int Count
        {
            get
            {
                return _list.Count;
            }
        }

        public string Top()
        {
            return _list.TailValue;
        }

        public void Clear()
        {
            _list.ClearList();
        }

        #endregion

        #region temp
        public void Print(System.Windows.Forms.ListBox listBox)
        {
            _list.PrintStackToListBox(listBox);
        }
        #endregion
    }
}
