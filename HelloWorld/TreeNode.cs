using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.Ifc4.ProductExtension;

namespace HelloWorld
{
    public class TreeNode<T> : INotifyPropertyChanged
    {
        public T Data { get; set; }

        public bool? IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                _isSelected = value;

                OnPropertyChanged(nameof(IsSelected));

                SyncChildValue();

                SyncParentValue();
            }
        }

/*        public bool IsVisible
        {
            get
            {
                return _isVisible;
            }
            set
            {
                _isVisible = value;

                OnPropertyChanged(nameof(IsVisible));
            }
        }*/

        private void SyncParentValue()
        {
            if (ParentNode == null)
            {
                return;
            }

            if (ParentNode.CheckIfChildrenAllTrue())
            {
                ParentNode.IsSelected = true;
            }
            else if (ParentNode.CheckIfChildrenAllFalse())
            {
                ParentNode.IsSelected = false;
            }
            else
            {
                ParentNode.IsSelected = null;
            }
        }

        private bool? _isSelected = false;

/*        private bool _isVisible = true;*/

        public List<TreeNode<T>> Children { get; set; } = new List<TreeNode<T>>();

        public TreeNode<T> ParentNode { get; set; }

        public TreeNode(T data)
        {
            Data = data;
        }

        public IfcElement Element
        {
            get
            {
                return _element;
            }
            set
            {
                _element = value;
            }
        }

        private IfcElement _element;

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void SyncChildValue()
        {
            if (Children.Count == 0)
            {
                return;
            }

            if (IsSelected == true)
            {
                foreach (var child in Children)
                {
                    if (child.IsSelected != true)
                    {
                        child.IsSelected = true;
                    }
                }
            }
            else if (IsSelected == false)
            {
                foreach (var child in Children)
                {
                    if (child.IsSelected != false)
                    {
                        child.IsSelected = false;
                    }
                }
            }
        }

        public bool CheckIfChildrenAllTrue()
        {
            foreach (var child in Children)
            {
                if (child.IsSelected != true)
                {
                    return false;
                }
            }

            return true;
        }

        public bool CheckIfChildrenAllFalse()
        {
            foreach (var child in Children)
            {
                if (child.IsSelected != false)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
