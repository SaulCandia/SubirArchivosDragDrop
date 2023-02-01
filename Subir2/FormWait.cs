using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Subir2
{
    public partial class FormWait : Form
    {

        public Action aSyncLoading { get; set; }
        public FormWait(Action loading)
        {
            InitializeComponent();
            if (loading == null)
                throw new ArgumentNullException();
            aSyncLoading = loading;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Task.Factory.StartNew(aSyncLoading).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}
