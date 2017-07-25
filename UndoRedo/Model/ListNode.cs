using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UndoRedo.Model
{
    class ListNode
    {
        private string _stringValue;
        private int _intValue;
        private ListNode _nextNode;
        private ListNode _previousNode;

        public ListNode(string value, int index, ListNode nextNode, ListNode previousNode)
        {
            _stringValue = value;
            _intValue = index;
            _nextNode = nextNode;
            _previousNode = previousNode;
        }

        public ListNode NextNode
        {
            get
            {
                return _nextNode;
            }
            set
            {
                _nextNode = value;
            }
        }

        public ListNode PreviousNode
        {
            get
            {
                return _previousNode;
            }
            set
            {
                _previousNode = value;
            }
        }

        public string Value
        {
            get
            {
                return _stringValue;
            }
        }

        public int IntValue
        {
            get
            {
                return _intValue;
            }
            set
            {
                _intValue = value;
            }
        }
    }
}
