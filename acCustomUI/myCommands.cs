// (C) Copyright 2020 by  
//
using System;
using System.Windows;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Windows;
using Autodesk.AutoCAD.Ribbon;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Windows;
using System.Drawing;
using System.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using System.Windows.Interop;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Internal.Windows;
using Autodesk.AutoCAD.Interop;
using Autodesk.Windows.ToolBars;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using MenuItem = Autodesk.AutoCAD.Windows.MenuItem;
using RibbonButton = Autodesk.Windows.RibbonButton;
using RibbonControl = Autodesk.Windows.RibbonControl;
using RibbonPanelSource = Autodesk.Windows.RibbonPanelSource;
using UserControl = System.Windows.Controls.UserControl;
using System.Reflection;
using Autodesk.AutoCAD.Windows.ToolPalette;
using Autodesk.Internal.InfoCenter;
using Autodesk.Private.InfoCenter;
using Visibility = System.Windows.Visibility;
using Window = Autodesk.AutoCAD.Windows.Window;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in_acCustomUI.MyCommands))]

namespace AutoCAD_CSharp_plug_in_acCustomUI
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here

        }

        // Modal Command with pickfirst selection
        [CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void MyPickFirst() // This method can have any name
        {
            PromptSelectionResult result = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection();
            if (result.Status == PromptStatus.OK)
            {
                // There are selected entities
                // Put your command using pickfirst set code here
            }
            else
            {
                // There are no selected entities
                // Put your command code here
            }
        }

        // Application Session Command with localized name
        [CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal | CommandFlags.Session)]
        public void MySessionCmd() // This method can have any name
        {
            // Put your command code here
        }

        // LispFunction is similar to CommandMethod but it creates a lisp 
        // callable function. Many return types are supported not just string
        // or integer.
        [LispFunction("MyLispFunction", "MyLispFunctionLocal")]
        public int MyLispFunction(ResultBuffer args) // This method can have any name
        {
            // Put your command code here

            // Return a value to the AutoCAD Lisp Interpreter
            return 1;
        }

        private static PaletteSet myPaletteSetLeft = null;

        [CommandMethod("ZsyPaletteSetLeft")]
        public static void ZsyPaletteSetLeft()
        {
            if (myPaletteSetLeft != null) return;
            myPaletteSetLeft = new PaletteSet("test PaletteSet left");
            var s = new System.Drawing.Size(100,800);
            myPaletteSetLeft.Size = s;
            myPaletteSetLeft.MinimumSize = new System.Drawing.Size(100,800) ;
            MyDockBarLeft x1 = new MyDockBarLeft();
            myPaletteSetLeft.Add("2", x1);
            myPaletteSetLeft.Dock = DockSides.Left;
            myPaletteSetLeft.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetLeft.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
        }

        private static PaletteSet myPaletteSetRight = null;
        [CommandMethod("ZsyPaletteSetRight")]
        public static void ZsyPaletteSetRight()
        {
            if (myPaletteSetRight != null) return;
            myPaletteSetRight = new PaletteSet("test PaletteSet right");
            var s = new System.Drawing.Size(200, 800);
            myPaletteSetRight.Size = s;
            myPaletteSetRight.MinimumSize = new System.Drawing.Size(300, 800);
            MyDockBarRight x1 = new MyDockBarRight();
            myPaletteSetRight.Add("3", x1);
            
            myPaletteSetRight.TitleBarLocation = PaletteSetTitleBarLocation.Right;
            myPaletteSetRight.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetRight.DockEnabled = DockSides.Right;
            myPaletteSetRight.Dock = DockSides.Right;
        }

        private static PaletteSet myPaletteSetTop = null;
        [CommandMethod("ZsyPaletteSetTop")]
        public static void ZsyPaletteSetTop()
        {
            if (myPaletteSetTop != null) return;
            myPaletteSetTop = new PaletteSet("test PaletteSet top");
            var s = new System.Drawing.Size(100, 100);
            myPaletteSetTop.Size = s;
            myPaletteSetTop.MinimumSize = s;
            MyDockBarTop x1 = new MyDockBarTop();
            myPaletteSetTop.Add("4", x1);
            myPaletteSetTop.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetTop.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetTop.Size = s;
            myPaletteSetTop.MinimumSize = s;
            myPaletteSetTop.DockEnabled = DockSides.Top;
            myPaletteSetTop.Dock = DockSides.Top; 
        }

        private static PaletteSet myPaletteSetBottom = null;
        [CommandMethod("ZsyPaletteSetBottom")]
        public static void ZsyPaletteSetBottom()
        {
            if (myPaletteSetBottom != null) return;
            myPaletteSetBottom = new PaletteSet("test PaletteSet bottom");
            var s = new System.Drawing.Size(50, 50);
            myPaletteSetBottom.Size = s;
            myPaletteSetBottom.MinimumSize = s;
            MyDockBarBottom x1 = new MyDockBarBottom();
            myPaletteSetBottom.Add("5", x1);
            
            myPaletteSetBottom.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetBottom.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetBottom.Size = s;
            myPaletteSetBottom.MinimumSize = s;
            myPaletteSetBottom.Dock = DockSides.Bottom;
            myPaletteSetBottom.DockEnabled = DockSides.Bottom;
        }

        [CommandMethod("ZsyMainMenu")]
        public static void ZsyMainMenu()
        {
            //获取CAD应用程序
            AcadApplication app = (AcadApplication)Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
            for (int i = 0; i < app.MenuGroups.Item(0).Menus.Count; i++)
            {
                app.MenuGroups.Item(0).Menus.RemoveMenuFromMenuBar(i);
            }

        }

        [CommandMethod("ZsyToolbar")]
        public static void ZsyToolbar()
        {
            Autodesk.AutoCAD.Interop.AcadApplication acadApp = (AcadApplication)Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
            Autodesk.AutoCAD.Interop.AcadToolbar atb = null;
            Autodesk.AutoCAD.Interop.AcadToolbarItem atbi = null;

            AcadMenuGroup mg = acadApp.MenuGroups.Item("ACAD");

            AcadToolbars atbs = mg.Toolbars;
            for (int i = 0; i < atbs.Count; i++)
            {
                atbs.Item(i).Visible = false;
            }

            try
            {
                String dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); 

                atb = atbs.Add("您自己工具栏的名称");
                //示例atbi = atb.AddToolbarButton(0, "工具栏名称", "帮助", "命令 ", false);
                //命令后，紧跟一空格，否则处于等待状态
                atbi = atb.AddToolbarButton(0, "数字文本求和", "数字文本求和", "Sum ", false);
                //示例atbi.SetBitmaps(工具栏图片文件路径,工具栏图片文件路径);
                atbi.SetBitmaps(dir + "\\toolbutton1.bmp", dir + "\\toolbutton1.bmp");

                atbi = atb.AddToolbarButton(1, "批量转换", "批量转换", "Sum ", false);
                atbi.SetBitmaps(dir + "\\toolbutton2.bmp", dir + "\\toolbutton2.bmp");

                atbi = atb.AddToolbarButton(2, "test", "test", "Sum ", false);
                atbi.SetBitmaps(dir + "\\toolbutton3.bmp", dir + "\\toolbutton3.bmp");
                
                //工具栏靠右边停靠
                atb.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockRight);
                atb.Visible = true;
            }
            catch(Exception ex)
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("加载工具栏出错……\n" + ex.Message);
            }
        }

        private static RibbonTab myRibbonTab = null;
        [CommandMethod("ZsyRibbonTab")]
        public static void ZsyRibbonTab()
        {
            // 在AutoCAD的Ribbon窗口中显示
            if (RibbonServices.RibbonPaletteSet == null)
            {
                Commands.Ribbon();
            }

            RibbonControl ribCntrl = RibbonServices.RibbonPaletteSet.RibbonControl;

            // 创建Ribbon Tab页
            if (myRibbonTab == null)
            {
                myRibbonTab = new RibbonTab();
                myRibbonTab.Title = "测试Ribbon页";
                myRibbonTab.Id = "MyRibbonTab";

                //
                RibbonPanel panel1 = new RibbonPanel();
                RibbonPanelSource panel1Src = new RibbonPanelSource();
                panel1Src.Title = "Panel 1";
                panel1.Source = panel1Src;
                myRibbonTab.Panels.Add(panel1);

                RibbonButton rbnBtnLine = NewRibbonBtn("line", "line ");
                panel1Src.Items.Add(rbnBtnLine);
                RibbonButton rbnBtnCircle = NewRibbonBtn("circle", "ZsyNewCircle ");
                panel1Src.Items.Add(rbnBtnCircle);

                //
                RibbonPanel panel2 = new RibbonPanel();
                RibbonPanelSource panel2Src = new RibbonPanelSource();
                panel2Src.Title = "Panel 2";
                panel2.Source = panel2Src;
                myRibbonTab.Panels.Add(panel2);

                RibbonButton rbnBtnTest2 = NewRibbonBtn2("画直线", "");
                panel2Src.Items.Add(rbnBtnTest2);

                RibbonButton rbnBtnTest3 = NewRibbonBtn3("画园", "");
                panel2Src.Items.Add(rbnBtnTest3);
            }

            bool isShow = false;
            foreach (RibbonTab rbnTab in ribCntrl.Tabs)
            {
                if (rbnTab == myRibbonTab)
                {
                    isShow = true;
                    break;
                }
                else
                {
                    rbnTab.IsVisible = false;
                }
            }
            if (!isShow)
            {
                ribCntrl.Tabs.Add(myRibbonTab);
            }
        }
        private static RibbonButton NewRibbonBtn(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;
            
            button.CommandParameter = cmd;
            return button;
        }
        private static RibbonButton NewRibbonBtn2(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;
            button.CommandHandler = new LineCommandHandler();
            button.CommandParameter = cmd;
            return button;
        }

        private static RibbonButton NewRibbonBtn3(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;
            button.CommandHandler = new CircleCommandHandler();
            button.CommandParameter = cmd;
            return button;
        }

       [CommandMethod("ZsyCmdTest")]
        public static void MyCmdTest()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("www.caxdev.com\n");
        }


       //右键菜单
       [CommandMethod("ZsyAddDefaultContextMenuItem")]
       public static void ZsyAddDefaultContextMenuItem()
       {
                //右键菜单 对象
                ContextMenuExtension m_ContextMenu = new ContextMenuExtension();
                m_ContextMenu.Title = "test 右键菜单";
                //右键菜单项及其事件
                MenuItem MenuItem_1 = new MenuItem("test 画圆");
                MenuItem_1.Click += new EventHandler(MenuItem_1_Click);

                //菜单项添加到右键菜单
                m_ContextMenu.MenuItems.Add(MenuItem_1);

                //加入到应用程序级的右键菜单中
                Application.AddDefaultContextMenuExtension(m_ContextMenu);
       }

       //选定右键菜单
       [CommandMethod("ZsyAddObjectContextMenuItem")]
       public static void ZsyAddObjectContextMenuItem()
       {
           //右键菜单 对象
           ContextMenuExtension m_ContextMenu = new ContextMenuExtension();
           m_ContextMenu.Title = "test 选定后右键菜单";
           //右键菜单项及其事件
           MenuItem MenuItem_1 = new MenuItem("test 画圆");
           MenuItem_1.Click += new EventHandler(MenuItem_1_Click);

           //菜单项添加到右键菜单
           m_ContextMenu.MenuItems.Add(MenuItem_1);

           //加入到某一种对象的右键菜单中
           Application.AddObjectContextMenuExtension(Entity.GetClass(typeof(Entity)), m_ContextMenu);
       }

       private static void MenuItem_1_Click(object sender, System.EventArgs e)
        {
            CircleCommandHandler.ZsyNewCircle();
        }


       [CommandMethod("ZsyCloseFileTabOnTop")]
       public static void ZsyCloseFileTabOnTop()
       {
           Commands.CloseFileTab();
       }

       [CommandMethod("ZsyCloseRibbon")]
       public static void ZsyCloseRibbon()
       {
           if (RibbonServices.RibbonPaletteSet != null)
           {
               Commands.RibbonClose();
           }
       }

        //修改当前工作空间:草图与注释
        [CommandMethod("ZsyWorkspace")]
       public static void ZsyWorkspace()
       {
            //"WS_Anno2DDraft"  "草图与注释"
            //"WS_3DModeling   "三维建模"
            //"WS_SM_0001"     "三维基础"
            //"WS_AcadClassic" "AutoCAD 经典"

            string curCuiFileName = Application.GetSystemVariable("MENUNAME").ToString() + ".cuix";
            CustomizationSection curCui = new CustomizationSection(curCuiFileName);

            
            string wsCurrentName = (string)Application.GetSystemVariable("WSCURRENT");

            foreach (Workspace curCuiWorkspace in curCui.Workspaces)
            {
                if ("WS_Anno2DDraft".Equals(curCuiWorkspace.ElementID) && !curCuiWorkspace.Name.Equals(wsCurrentName)) //
                {
                    Application.SetCurrentWorkspace(curCuiWorkspace.Name);
                    return;
                }
            }
        }

       [CommandMethod("ZsyPaletteSetClose")]
       public static void ZsyPaletteSetClose()
       {
          
       }
    }

}
