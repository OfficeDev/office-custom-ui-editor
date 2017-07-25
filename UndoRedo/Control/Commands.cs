using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UndoRedo.Model;

namespace UndoRedo.Control
{
    /// <summary>
    /// This class contains the logic for the commands.
    /// </summary>
    public class Commands
    {
        private UndoRedoStack _undoStack;
        private UndoRedoStack _redoStack;
        private string _lastUndoActionName;
        private string _lastSavedText;
        private string _lastRemovedRtf;
        private int _lastIndex;
        private bool _lastUndoActionWasBackspace;

        public Commands()
        {
            _undoStack = new UndoRedoStack();
            _undoStack.Push(string.Empty, 0);
            _redoStack = new UndoRedoStack();
            _lastUndoActionName = string.Empty;
            _lastSavedText = string.Empty;
            _lastRemovedRtf = string.Empty;
            _lastUndoActionWasBackspace = false;
        }

        public Commands(string initialRtf)
        {
            _undoStack = new UndoRedoStack();
            _undoStack.Push(initialRtf, 0);
            _redoStack = new UndoRedoStack();
            _lastUndoActionName = string.Empty;
            _lastSavedText = string.Empty;
            _lastRemovedRtf = string.Empty;
            _lastUndoActionWasBackspace = false;
        }

        public void NewCommand(string undoActionName, string rtf, string text, int index)
        {
            int popIndex;
            switch (undoActionName)
            {
                case "Typing":
                    if (!"Typing".Equals(_lastUndoActionName))
                    {
                        if (!string.IsNullOrEmpty(text) && _undoStack.Count == 1 && !string.IsNullOrEmpty(_lastUndoActionName))
                        {
                            //this is an special case where we have emptied the stack (because an undo), there is some text in the editor
                            //and we start typing again, then we save the current status (previously removed by undo) and
                            //then we start saving the new status.
                            if (!string.IsNullOrEmpty(_lastRemovedRtf)) _undoStack.Push(_lastRemovedRtf, _lastIndex);
                            _undoStack.Push(rtf, index);
                            _lastUndoActionName = undoActionName;
                        }
                        else
                        {
                            //normal case
                            _undoStack.Push(rtf, index);
                            _lastUndoActionName = undoActionName;
                        }
                    }
                    else
                    {
                        //here we have 2 cases (one for typing and one for backspace)
                        if (text.Length > _lastSavedText.Length)
                        {
                            //if (!_undoStack.IsEmpty()) _undoStack.Pop(out popIndex);
                            _undoStack.Push(rtf, index);
                            _lastUndoActionWasBackspace = false;
                        }
                        else
                        {
                            _undoStack.Push(rtf, index);
                            _lastUndoActionWasBackspace = true;
                        }
                    }
                    _redoStack.Clear();
                    _lastSavedText = text;
                    break;
                case "Cut":
                case "Paste":
                case "Delete":
                    UndoStackManipulation(undoActionName, rtf, text, index, out popIndex);
                    break;
            }
        }

        /// <summary>
        /// Basic operation for managing the different commands.
        /// </summary>
        /// <param name="undoActionName">Name identifier for the current undo action.</param>
        /// <param name="rtf">Current RTF in the RichTextBox control.</param>
        /// <param name="text">Current Text in the RichTextBox control.</param>
        /// <param name="index">Current Index in the RichTextBox control.</param>
        /// <param name="popIndex">The index to be used as the SelectionStart attribute.</param>
        private void UndoStackManipulation(string undoActionName, string rtf, string text, int index, out int popIndex)
        {
            popIndex = 0;
            if (!undoActionName.Equals(_lastUndoActionName))
            {
                if (!string.IsNullOrEmpty(text) && _undoStack.Count == 1 && !string.IsNullOrEmpty(_lastUndoActionName))
                {
                    if (!string.IsNullOrEmpty(_lastRemovedRtf)) _undoStack.Push(_lastRemovedRtf, _lastIndex);
                    _undoStack.Push(rtf, index);
                    _lastUndoActionName = undoActionName;
                }
                else
                {
                    _undoStack.Push(rtf, index);
                    _lastUndoActionName = undoActionName;
                }
            }
            else
            {
                //if (!_undoStack.IsEmpty()) _undoStack.Pop(out popIndex);
                _undoStack.Push(rtf, index);
            }
            _lastSavedText = text;
            _redoStack.Clear();
        }

        /// <summary>
        /// Manages the Undo Stack.
        /// </summary>
        /// <param name="index">The index to be used as the SelectionStart attribute.</param>
        /// <param name="canRedo">Identifies if the command can be redone.</param>
        /// <returns>RTF string.</returns>
        public string RemoveCommand(out int index, bool canRedo)
        {
            string removedRtf = _undoStack.Pop(out index);

            if (_lastUndoActionName.Equals("Undo") && canRedo && !string.IsNullOrEmpty(_lastRemovedRtf) && !_undoStack.IsEmpty()
                && !_redoStack.Top().Equals(_lastRemovedRtf))
            {
                _redoStack.Push(_lastRemovedRtf, _lastIndex);
            }
            else if (canRedo && !string.IsNullOrEmpty(removedRtf.Trim()))
            {
                _redoStack.Push(removedRtf, index);
            }
            else if(!string.IsNullOrEmpty(removedRtf.Trim()))
            {
                _redoStack.Push(removedRtf, index);
            }

            if (_undoStack.IsEmpty())
            {
                _undoStack.Push(string.Empty, 0);
            }

            _lastUndoActionName = "Undo";
            _lastRemovedRtf = removedRtf;
            _lastIndex = index;
            return _lastRemovedRtf;
        }

        /// <summary>
        /// Manages the Redo Stack.
        /// </summary>
        /// <param name="index">The index to be used as the SelectionStart attribute.</param>
        /// <returns>RTF string.</returns>
        public string RedoCommand(out int index)
        {
            if(_redoStack.IsEmpty())
            {
                index = 0;
                return null;
            }

            string rtf = _redoStack.Pop(out index);

            //avoid repeating the insertion if _lastRemovedRtf when Redo is clicked multiple times
            if (!_lastUndoActionName.Equals("Redo") && !string.IsNullOrEmpty(_lastRemovedRtf) && !_undoStack.Top().Equals(_lastRemovedRtf))
            {
                _undoStack.Push(_lastRemovedRtf, _lastIndex);
            }

            //avoid inserting the same value two times to the top of the stack
            if (!_undoStack.Top().Equals(rtf)) _undoStack.Push(rtf, index);

            _lastUndoActionName = "Redo";

            return rtf;
        }

        public void PrintStack(System.Windows.Forms.ListBox listBox, bool isUndoStack)
        {
            if (isUndoStack)
                _undoStack.Print(listBox);
            else
                _redoStack.Print(listBox);
        }
    }
}
