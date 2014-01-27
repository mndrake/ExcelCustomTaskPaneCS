namespace ExcelCustomTaskPaneCS
{
    using System;
    using System.Windows.Forms;
    using System.Windows.Forms.Integration;

    using ExcelDna.Integration;
    using ExcelDna.Integration.CustomUI;


    /////////////// Define the backing class for the Ribbon ///////////////////////////
    // Would need to be marked with [ComVisible(true)] if in a project that is marked as [assembly:ComVisible(false)] which is the default for VS projects.
    public class MyRibbon : ExcelRibbon
    {
        public void OnShowCTP(IRibbonControl control)
        {
            CTPManager.ShowCTP();
        }


        public void OnDeleteCTP(IRibbonControl control)
        {
            CTPManager.DeleteCTP();
        }
    }

    /////////////// Define the UserControl to display on the CTP ///////////////////////////
    // Would need to be marked with [ComVisible(true)] if in a project that is marked as [assembly:ComVisible(false)] which is the default for VS projects.
    public class MyUserControl : UserControl
    {
//        public Label TheLabel;
        public TaskPaneContent Content;

        public MyUserControl()
        {
            Content = new TaskPaneContent();
            var wpfElementHost = new ElementHost() { Dock = DockStyle.Fill };
            wpfElementHost.HostContainer.Children.Add(Content);
            Content.MyButton.Click += MyButton_Click;
            
            Controls.Add(wpfElementHost);
        }

        void MyButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MessageBox.Show("You clicked the button.");
        }
    }

    /////////////// Helper class to manage CTP ///////////////////////////
    internal static class CTPManager
    {
        static CustomTaskPane ctp;

        public static void ShowCTP()
        {
            if (ctp == null)
            {
                // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(MyUserControl), "My Super Task Pane");
                ctp.Visible = true;
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                ctp.DockPositionStateChange += ctp_DockPositionStateChange;
                ctp.VisibleStateChange += ctp_VisibleStateChange;
            }
            else
            {
                // Just show it again
                ctp.Visible = true;
            }
        }


        public static void DeleteCTP()
        {
            if (ctp != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctp.Delete();
                ctp = null;
            }
        }

        static void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            MessageBox.Show("Visibility changed to " + CustomTaskPaneInst.Visible);
        }

        static void ctp_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            ((MyUserControl)ctp.ContentControl).Content.MyLabel.Content = "Moved to " + CustomTaskPaneInst.DockPosition.ToString();
        }
    }
}
