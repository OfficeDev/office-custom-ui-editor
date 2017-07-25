using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UndoRedo.Model
{
    public class List
    {
        private ListNode _head;
        private ListNode _tail;
        private int _nodeCount;

        public List()
        {
            _head = null;
            _tail = null;
            _nodeCount = 0;
        }

        public void AddToTail(string value, int index)
        {
            if (_tail != null)
            {
                _tail.NextNode = new ListNode(value, index, null, _tail);
                _tail = _tail.NextNode;
            }
            else
            {
                _tail = new ListNode(value, index, null, null);
                _head = _tail;
            }

            ++_nodeCount;
        }

        public void AddToHead(string value, int index)
        {
            if (!IsEmpty())
            {
                ListNode tmp = new ListNode(value, index, _head, null);
                _head = tmp.NextNode;
            }
            else
            {
                _head = new ListNode(value, index, null, null);
                _tail = _head;
            }

            ++_nodeCount;
        }

        public string RemoveFromTail(out int index)
        {
            if (!IsEmpty())
            {
                string value = _tail.Value;
                index = _tail.IntValue;

                if (_head == _tail)
                {
                    _head = null;
                    _tail = null;
                }
                else
                {
                    _tail = _tail.PreviousNode;
                    _tail.NextNode = null;
                }

                --_nodeCount;
                return value;
            }
            else
            {
                throw new AccessViolationException("The List is emtpy.");
            }
        }

        public string RemoveFromHead(out int index)
        {
            if (!IsEmpty())
            {
                string value = _head.Value;
                index = _head.IntValue;

                if (_head == _tail)
                {
                    _head = null;
                    _tail = null;
                }
                else
                {
                    _head = _head.NextNode;
                    _head.PreviousNode = null;
                }

                --_nodeCount;
                return value;
            }
            else
            {
                throw new AccessViolationException("The List is emtpy.");
            }
        }

        public bool IsEmpty()
        {
            return _nodeCount == 0;
        }

        public string TailValue
        {
            get
            {
                return _tail.Value;
            }
        }

        public int Count
        {
            get
            {
                return _nodeCount;
            }
        }

        public void ClearList()
        {
            _head = null;
            _tail = null;
            _nodeCount = 0;
        }


        #region temp
        public void PrintStackToListBox(System.Windows.Forms.ListBox listBox)
        {
            listBox.Items.Clear();
            for (ListNode tmp = _head; tmp != null; tmp = tmp.NextNode)
            {
                listBox.Items.Add(tmp.Value);
            }
        }
        #endregion
    }
}
