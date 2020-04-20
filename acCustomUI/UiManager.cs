// (C) Copyright 2020 by  
//
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
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
using System.IO;
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
using System.Windows.Input;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.Internal.Reactors;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Windows.ToolPalette;
using Autodesk.Internal.InfoCenter;
using Autodesk.Private.InfoCenter;
using Microsoft.Win32;
using Visibility = System.Windows.Visibility;
using Window = Autodesk.AutoCAD.Windows.Window;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in_acCustomUI.UiManager))]

namespace AutoCAD_CSharp_plug_in_acCustomUI
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!

    public class UiManager
    {
        private const string REG_VERSION_SUB_KEY_HKLM = @"R23.1\ACAD-3001";
        private const string REG_VERSION_SUB_KEY_ZHCN = @"R23.1\ACAD-3001:804";
        private const string PROFILE_MC2020 = "MCSL2020";
        private const string MG_NAME_MCMENU = "MCSL2020MENU";
        private const string McWorkspaceName = "AutoCAD 经典";

        private static AcadApplication AcadApp = Application.AcadApplication as AcadApplication;

        private static void McWriteMessage(string message)
        {
            if (Application.DocumentManager.MdiActiveDocument != null && Application.DocumentManager.MdiActiveDocument.Editor != null)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(message);
            }
        }


        private static bool get_is_mcsl2020_environment()
        {
            foreach (string arg in Environment.GetCommandLineArgs())
            {
                if (arg.Contains(PROFILE_MC2020.ToLower()) || arg.Contains(PROFILE_MC2020.ToUpper()))
                    return true;
            }

            return false;
        }

        [CommandMethod("McUiSwitch")]
        public void McUiSwitch()
        {
            if (mcUiLoaded)
            {
                McUnLoad(false);
            }
            else
            {
                McLoad(false);
            }
        }

        private static bool mcUiLoaded = false;

        [CommandMethod("McLoad")]
        public void McLoad()
        {
            McLoad(false);
        }

        [CommandMethod("McLoad")]
        public void McLoad(bool initPlugin)
        {
            try
            {
                McWriteMessage("MyPlugin 插件 UI 加载处理开始\n");

                if (initPlugin)
                {
                    IAcadPreferences acadPreferences = (IAcadPreferences) Application.Preferences;

                    if (!PROFILE_MC2020.Equals(acadPreferences.Profiles.ActiveProfile,
                        StringComparison.CurrentCultureIgnoreCase)) //当前配置不是 MCSL2020 的话
                    {
                        McWriteMessage("当前配置不是 MCSL2020，说明不是启动算量软件，不对界面做处理");
                        return;
                    }
                }

                Application.UserConfigurationManager.CurrentProfileChanged += new ProfileEventHandler((sender, e) =>
                {
                    McWriteMessage("CurrentProfileChanged : ProfileName: " + e.ProfileName + "\n");
                });


                AddAppEvent();
                McWriteMessage("0.0  AddAppEvent()\n");

                //使用不同的配置文件，与 AutoCAD 独立启动的情况区分开
                McSetActiveProfile();
                McWriteMessage("0.1  McSetActiveProfile()\n");

                //Cui
                //使用starter 时不需要执行
                //McCustomCui();
                //McWriteMessage("0.2  McCustomCui()\n");

                //把工作空间设置为 “AutoCAD 经典”,这步必须最先处理
                McWorkspace();
                McWriteMessage("0.3  McWorkspace()\n");

                //无选定对象时的右键菜单
                McAddDefaultContextMenuItem();
                McWriteMessage("1 完成 McAddDefaultContextMenuItem()\n");

                //有选定对象时的右键菜单
                McAddObjectContextMenuItem();
                McWriteMessage("2  McAddObjectContextMenuItem()\n");


                //顶部Ribbon 菜单区域，AutoCAD自带tab隐藏，添加自定义tab、按钮
                McRibbonTab();
                McWriteMessage("3  McRibbonTab()\n");

                //TODO 关闭自带的所有 palette 是否有必要？
                //McPaletteSetClose();

                //上下左右添加自定义停靠区域
                McPaletteSetLeft(); //包含弹出模态窗口、非模态置顶窗口
                McWriteMessage("4  McPaletteSetLeft()\n");
                McPaletteSetTop();
                McWriteMessage("5  McPaletteSetTop()\n");
                McPaletteSetBottom();
                McWriteMessage("6  McPaletteSetBottom()\n");
                McPaletteSetRight(); //WPF UserControl
                McWriteMessage("7  McPaletteSetRight()\n");

                //监听选定事件，选定对象id 在 非模态置顶窗口中显示
                MyEvent.AddSelectChangeEvent();
                McWriteMessage("8  MyEvent.AddSelectChangeEvent()\n");

                //隐藏AutoCAD自带 toolbar
                //McHideAllToolbar();
                //McWriteMessage("9.1  McHideAllToolbar()\n");

                //添加自定义toolbar
                McCustomToolbar();
                McWriteMessage("9.2  McCustomToolbar()\n");

                //隐藏绘图区域顶部的“文件选项卡”tab
                McCloseFileTabOnTop();
                McWriteMessage("10  McCloseFileTabOnTop()\n");

                //彻底隐藏 Robbon 菜单
                McCloseRibbon();
                McWriteMessage("11  McCloseRibbon()\n");

                //主菜单隐藏原有部分
                McCustomMainMenuClear();
                McWriteMessage("12  McCustomMainMenuClear()\n");

                McWriteMessage("13  McCustomMainMenuBizShow()\n");
                McCustomMainMenuBizShow();

                mcUiLoaded = true;

                McWriteMessage("MyPlugin 插件 UI 已加载\n");
            }
            catch (System.Exception ex)
            {
                if (ex is Exception)
                {
                    McWriteMessage("AutoCAD runtime Exception\n" + ex.Message + "\n");
                }
                else
                {
                    McWriteMessage(".NET System Exception\n" + ex.Message + "\n");
                }
            }
            finally
            {
                resetProfileInRegistry(); //注册表中记录的值重置为默认配置文件，AutoCAD 独立启动时仍可使用默认配置文件
                McWriteMessage("注册表选定配置信息已重置\n");
            }
        }

        [CommandMethod("McUnLoad")]
        public void McUnLoad()
        {
            McUnLoad(false);
        }

        public void McUnLoad(bool terminate)
        {
            if (!terminate)
            {
                McWriteMessage("MyPlugin 插件 UI 卸载开始\n");

                //McPaletteSetLeftUn();
                //McWriteMessage("1  McPaletteSetLeftUn()\n");
                McPaletteSetTopUn();
                McWriteMessage("2  McPaletteSetTopUn()\n");
                McPaletteSetBottomUn();
                McWriteMessage("3  McPaletteSetBottomUn()\n");
                McPaletteSetRightUn(); //WPF UserControl
                McWriteMessage("4  McPaletteSetRightUn()\n");

                McSetActiveProfileUn();
                McWriteMessage("5  McSetActiveProfileUn()\n");

                McCustomMainMenuSwitchShow();
                McWriteMessage("6  McCustomMainMenuSwitchShow()\n");

                mcUiLoaded = false;

                McWriteMessage("MyPlugin 插件 UI 已卸载\n");
            }
            else
            {
                //TODO
            }
        }

        public void McUnLoad____(bool terminate)
        {
            try
            {
                McWriteMessage("MyPlugin 插件 UI 卸载开始\n");

                //把工作空间设置为 “AutoCAD 经典”,这步必须最先处理
                McWorkspaceUn();
                McWriteMessage("0  McWorkspaceUn()\n");

                //无选定对象时的右键菜单
                McAddDefaultContextMenuItemUn();
                McWriteMessage("1 完成 McAddDefaultContextMenuItemUn()\n");

                //有选定对象时的右键菜单
                McAddObjectContextMenuItemUn();
                McWriteMessage("2  McAddObjectContextMenuItemUn()\n");


                //顶部Ribbon 菜单区域，AutoCAD自带tab隐藏，添加自定义tab、按钮
                //McRibbonTabUn();
                //McWriteMessage("3  McRibbonTabUn()\n");

                //TODO 关闭自带的所有 palette 是否有必要？
                //McPaletteSetClose();

                //上下左右添加自定义停靠区域
                if (!terminate)
                {
                    McPaletteSetLeftUn();
                    McWriteMessage("4  McPaletteSetLeftUn()\n");
                    McPaletteSetTopUn();
                    McWriteMessage("5  McPaletteSetTopUn()\n");
                    McPaletteSetBottomUn();
                    McWriteMessage("6  McPaletteSetBottomUn()\n");
                    McPaletteSetRightUn(); //WPF UserControl
                    McWriteMessage("7  McPaletteSetRightUn()\n");

                    //TODO
                    //MyEvent.AddSelectChangeEventUn();
                    //McWriteMessage("8  MyEvent.AddSelectChangeEventUn()\n");
                }

                //隐藏AutoCAD自带 toolbar
                McHideAllToolbarUn();
                McWriteMessage("9.1  McHideAllToolbarUn()\n");

                //添加自定义toolbar
                if (!terminate)
                {
                    McCustomToolbarUn();
                    McWriteMessage("9.2  McCustomToolbarUn()\n");
                }

                //隐藏绘图区域顶部的“文件选项卡”tab
                McCloseFileTabOnTopUn();
                McWriteMessage("10  McCloseFileTabOnTopUn()\n");


                //彻底隐藏 Robbon 菜单
                McCloseRibbonUn();
                McWriteMessage("11  McCloseRibbonUn()\n");

                McWriteMessage("12  McCustomMainMenuBizHide()\n");
                McCustomMainMenuBizHide();

                McWriteMessage("13  McCustomMainMenuSwitchShow()\n");
                McCustomMainMenuSwitchShow();

                mcUiLoaded = false;

                McWriteMessage("MyPlugin 插件 UI 已卸载\n");
            }
            catch (System.Exception ex)
            {
                if (ex is Exception)
                {
                    McWriteMessage("AutoCAD runtime Exception\n" + ex.Message + "\n");
                }
                else
                {
                    McWriteMessage(".NET System Exception\n" + ex.Message + "\n");
                }
            }
        }


        //=============================================================================================================

        [CommandMethod("AddAppEvent")]
        public void AddAppEvent()
        {
            Application.SystemVariableChanged +=
                new Autodesk.AutoCAD.ApplicationServices.
                    SystemVariableChangedEventHandler(appSysVarChanged);
        }

        [CommandMethod("RemoveAppEvent")]
        public void RemoveAppEvent()
        {
            Application.SystemVariableChanged -=
                new Autodesk.AutoCAD.ApplicationServices.
                    SystemVariableChangedEventHandler(appSysVarChanged);
        }

        public void appSysVarChanged(object senderObj,
            Autodesk.AutoCAD.ApplicationServices.
                SystemVariableChangedEventArgs sysVarChEvtArgs)
        {
            object oVal = Application.GetSystemVariable(sysVarChEvtArgs.Name);

            // Display a message box with the system variable name and the new value
            McWriteMessage(sysVarChEvtArgs.Name + " was changed." +
                           "\nNew value: " + oVal.ToString() + "\n");

            //if ("WSCURRENT".Equals(sysVarChEvtArgs.Name))
            //{
            //    if (mcUiLoaded)
            //    {
            //        McCustomMainMenuClear();
            //        McCustomMainMenuBizShow();
            //        McCustomToolbar();
            //    }
            //    else
            //    {
            //        McCustomMainMenuSwitchShow();
            //    }

            //    McWriteMessage("工作区切换，菜单栏和工具栏要重新处理\n");
            //}
        }

        #region 原始示例代码

        /*
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
        */
        #endregion

        //=============================================================================================================

        private static PaletteSet myPaletteSetLeft = null;
        private static MyDockBarLeft dockBarLeft = null;
        [CommandMethod("McPaletteSetLeft")]
        public void McPaletteSetLeft()
        {
            if (myPaletteSetLeft != null)
            {
                myPaletteSetLeft.Visible = true;
                return;
            }
            myPaletteSetLeft = new PaletteSet("test dock left");
            var s = new System.Drawing.Size(130,800);
            myPaletteSetLeft.Size = s;
            myPaletteSetLeft.MinimumSize = new System.Drawing.Size(130,800) ;
            dockBarLeft = new MyDockBarLeft();
            myPaletteSetLeft.Add("2", dockBarLeft);
            
            myPaletteSetLeft.Style = 0;
            myPaletteSetLeft.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetLeft.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetLeft.Dock = DockSides.Left;
            myPaletteSetLeft.DockEnabled = DockSides.Left;
        }

        [CommandMethod("McPaletteSetLeftUn")]
        public void McPaletteSetLeftUn()
        {
            if (myPaletteSetLeft != null)
            {
                myPaletteSetLeft.Close();
                myPaletteSetLeft = null;
            }
            if (dockBarLeft != null)
            {
                dockBarLeft = null;
            }
        }

        private static PaletteSet myPaletteSetRight = null;
        private static MyDockBarRight dockBarRight = null;
        [CommandMethod("McPaletteSetRight")]
        public  void McPaletteSetRight()
        {
            if (myPaletteSetRight != null)
            {
                myPaletteSetRight.Visible = true;
                return;
            }
            myPaletteSetRight = new PaletteSet("test PaletteSet right");
            var s = new System.Drawing.Size(200, 800);
            myPaletteSetRight.Size = s;
            myPaletteSetRight.MinimumSize = new System.Drawing.Size(300, 800);
            dockBarRight = new MyDockBarRight();
            myPaletteSetRight.Add("3", dockBarRight);
            
            myPaletteSetRight.TitleBarLocation = PaletteSetTitleBarLocation.Right;
            myPaletteSetRight.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetRight.DockEnabled = DockSides.Right;
            myPaletteSetRight.Dock = DockSides.Right;
        }

        [CommandMethod("McPaletteSetRightUn")]
        public void McPaletteSetRightUn()
        {
            if (myPaletteSetRight != null)
            {
                myPaletteSetRight.Close();
                myPaletteSetRight = null;
            }
            if (dockBarRight != null)
            {
                dockBarRight = null;
            }
        }

        private static PaletteSet myPaletteSetTop = null;
        private static MyDockBarTop dockBarTop = null;
        [CommandMethod("McPaletteSetTop")]
        public void McPaletteSetTop()
        {
            if (myPaletteSetTop != null)
            {
                myPaletteSetTop.Visible = true;
                return;
            }
            myPaletteSetTop = new PaletteSet("test PaletteSet top");
            var s = new System.Drawing.Size(1920, 100);
            myPaletteSetTop.Size = s;
            myPaletteSetTop.MinimumSize = s;
            dockBarTop = new MyDockBarTop();
            myPaletteSetTop.Add("4", dockBarTop);
            myPaletteSetTop.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetTop.Style = PaletteSetStyles.NoTitleBar;
            
            myPaletteSetTop.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetTop.Size = s;
            myPaletteSetTop.MinimumSize = s;
            myPaletteSetTop.DockEnabled = DockSides.Top;
            myPaletteSetTop.Dock = DockSides.Top;
            myPaletteSetTop.Style = PaletteSetStyles.NoTitleBar;

            myPaletteSetTop.RolledUp = false;
            myPaletteSetTop.SizeChanged += new PaletteSetSizeEventHandler((sender, e) =>
            {
                if (e.Height != 100)
                {
                    myPaletteSetTop.Size = new System.Drawing.Size(1920,100);
                    myPaletteSetTop.Dock = DockSides.Top;
                }
                    
            });
        }

        [CommandMethod("McPaletteSetTopUn")]
        public void McPaletteSetTopUn()
        {
            if (myPaletteSetTop != null)
            {
                myPaletteSetTop.Close();
                myPaletteSetTop = null;
            }
            if (dockBarTop != null)
            {
                dockBarTop = null;
            }
        }

        private static PaletteSet myPaletteSetBottom = null;
        private static MyDockBarBottom dockBarBottom = null;
        [CommandMethod("McPaletteSetBottom")]
        public void McPaletteSetBottom()
        {
            if (myPaletteSetBottom != null)
            {
                myPaletteSetBottom.Visible = true;
                return;
            }
            myPaletteSetBottom = new PaletteSet("test PaletteSet bottom");
            var s = new System.Drawing.Size(50, 50);
            myPaletteSetBottom.Size = s;
            myPaletteSetBottom.MinimumSize = s;
            dockBarBottom = new MyDockBarBottom();
            myPaletteSetBottom.Add("5", dockBarBottom);
            
            myPaletteSetBottom.TitleBarLocation = PaletteSetTitleBarLocation.Left;
            myPaletteSetBottom.Visible = true;

            // Visible 设置之后再设置尺寸和停靠
            myPaletteSetBottom.Size = s;
            myPaletteSetBottom.MinimumSize = s;
            myPaletteSetBottom.Dock = DockSides.Bottom;
            myPaletteSetBottom.DockEnabled = DockSides.Bottom;
        }
        [CommandMethod("McPaletteSetBottomUn")]
        public void McPaletteSetBottomUn()
        {
            if (myPaletteSetBottom != null)
            {
                myPaletteSetBottom.Close();
                myPaletteSetBottom = null;
            }
            if (dockBarBottom != null)
            {
                dockBarBottom = null;
            }
        }

        [CommandMethod("MyTopWindowShow")]
        public void MyTopWindowShow()
        {
            MyTopWindow.showModeless();
        }

        [CommandMethod("McCustomToolbar")]
        public void McCustomToolbar()
        {
            string mcToolbarName = "MC工具栏";
            AcadMenuGroup mg = AcadApp.MenuGroups.Item(0);
            AcadToolbars atbs = mg.Toolbars;
            for (int i = 0; i < atbs.Count; i++)
            {
                if (mcToolbarName.Equals(atbs.Item(i).Name))
                {
                    atbs.Item(i).Visible = true;
                    return;
                }
            }

            String dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            
            AcadToolbar atbCustom = atbs.Add(mcToolbarName);
            
            //示例atbi = atbCustom.AddToolbarButton(0, "工具栏名称", "帮助", "命令 ", false);
            //命令后，紧跟一空格，否则处于等待状态
            AcadToolbarItem atbi = atbCustom.AddToolbarButton(0, "对象信息窗口", "Test 显示置顶非模态窗口", "MyTopWindowShow ", false);
            //示例atbi.SetBitmaps(工具栏图片文件路径,工具栏图片文件路径);
            atbi.SetBitmaps(dir + "\\toolbutton1.bmp", dir + "\\toolbutton1.bmp");
            
            atbi = atbCustom.AddToolbarButton(1, "画圆", "Test 画圆", "McNewCircle ", false);
            atbi.SetBitmaps(dir + "\\toolbutton2.bmp", dir + "\\toolbutton2.bmp");

            atbi = atbCustom.AddToolbarButton(2, "Test3", "Test3", "Sum ", false);
            atbi.SetBitmaps(dir + "\\toolbutton3.bmp", dir + "\\toolbutton3.bmp");

            atbi = atbCustom.AddToolbarButton(3, "Test4", "Test4", "Sum ", false);
            atbi.SetBitmaps(dir + "\\toolbutton4.bmp", dir + "\\toolbutton4.bmp");

            //工具栏靠右边停靠
            atbCustom.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockRight);
            atbCustom.Visible = true;
        }

        [CommandMethod("McCustomToolbarUn")]
        public void McCustomToolbarUn()
        {
            string mcToolbarName = "MC工具栏";
            AcadMenuGroup mg = AcadApp.MenuGroups.Item(0);

            AcadToolbars atbs = mg.Toolbars;
            for (int i = 0; i < atbs.Count; i++)
            {
                if (mcToolbarName.Equals(atbs.Item(i).Name))
                {
                    atbs.Item(i).Visible = false;
                    return;
                }
            }
        }

        //private static List<AcadToolbar> hiddenToobars = new List<AcadToolbar>();
        [CommandMethod("McHideAllToolbar")]
        public void McHideAllToolbar()
        {
            AcadMenuGroup mg = AcadApp.MenuGroups.Item(0);

            AcadToolbars atbs = mg.Toolbars;
            for (int i = 0; i < atbs.Count; i++)
            {
                if (atbs.Item(i).Visible)
                {
                    //hiddenToobars.Add(atbs.Item(i));
                    atbs.Item(i).Visible = false;
                }
            }
        }

        [CommandMethod("McHideAllToolbarUn")]
        public void McHideAllToolbarUn()
        {
            //if (hiddenToobars.Count > 0)
            //{
            //    hiddenToobars.ForEach((a) => a.Visible = true);
            //}
        }

        private static RibbonTab myRibbonTab = null;
        [CommandMethod("McRibbonTab")]
        public void McRibbonTab()
        {
            // 在AutoCAD的Ribbon窗口中显示
            if (RibbonServices.RibbonPaletteSet == null)
            {
                //Commands.Ribbon();
                return;
            }

            RibbonControl ribCntrl = RibbonServices.RibbonPaletteSet.RibbonControl;

            // 创建Ribbon Tab页
            if (myRibbonTab == null)
            {
                myRibbonTab = new RibbonTab();
                myRibbonTab.Title = "梦诚菜单";
                myRibbonTab.Id = "MyRibbonTabMengCheng";

                //
                RibbonPanel panel1 = new RibbonPanel();
                RibbonPanelSource panel1Src = new RibbonPanelSource();
                panel1Src.Title = "菜单";
                panel1.Source = panel1Src;
                myRibbonTab.Panels.Add(panel1);

                RibbonButton rbnBtnSwitch = NewRibbonBtn("CAD 界面切换", "McUiSwitch ");
                panel1Src.Items.Add(rbnBtnSwitch);

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
                    //其他autoCAD 自带的 ribbon 菜单, 不做改变
                }
            }
            if (!isShow)
            {
                ribCntrl.Tabs.Add(myRibbonTab);
            }
        }

        private RibbonButton NewRibbonBtn(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;

            button.CommandHandler = button.CommandHandler = new McUiSwitchCommand(this);

            button.CommandParameter = cmd;
            return button;
        }

        class McUiSwitchCommand : ICommand 
        {
            private readonly UiManager uiManager;

            public McUiSwitchCommand(UiManager uiManager)
            {
                this.uiManager = uiManager;
            }

            public void Execute(object parameter)
            {
                uiManager.McUiSwitch();
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged;
        }
        private RibbonButton NewRibbonBtn2(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;
            button.CommandHandler = new LineCommandHandler();
            button.CommandParameter = cmd;
            return button;
        }

        private RibbonButton NewRibbonBtn3(string text, string cmd)
        {
            RibbonButton button = new RibbonButton();
            button.Text = text;
            button.ShowText = true;
            button.CommandHandler = new CircleCommandHandler();
            button.CommandParameter = cmd;
            return button;
        }

        [CommandMethod("McRibbonTabUn")]
        public void McRibbonTabUn()
        {
            myRibbonTab = null;
        }

        //[CommandMethod("McCmdTest")]
        //public void MyCmdTest()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    doc.McWriteMessage("www.caxdev.com\n");
        //}

        [CommandMethod("McNewLine")]
        public void McNewLine()
        {
            LineCommandHandler.NewLine();
        }

        [CommandMethod("McNewCircle")]
        public void McNewCircle()
        {
            CircleCommandHandler.NewCircle();
        }

        private static void MenuItem_1_Click(object sender, System.EventArgs e)
        {
            CircleCommandHandler.NewCircle();
        }

        private static ContextMenuExtension defaultContextMenu = null;
        private static MenuItem defaultMenuItem_1 = null;
        private static EventHandler defaultMenuItem_1_Event = new EventHandler(MenuItem_1_Click);

        //右键菜单
        [CommandMethod("McAddDefaultContextMenuItem")]
        public void McAddDefaultContextMenuItem()
        {
            if (defaultContextMenu != null) { return;}

            //右键菜单 对象
            defaultContextMenu = new ContextMenuExtension();
            defaultContextMenu.Title = "Test 右键菜单";
            //右键菜单项及其事件
            defaultMenuItem_1 = new MenuItem("Test 画圆");
            defaultMenuItem_1.Click += defaultMenuItem_1_Event;

            //菜单项添加到右键菜单
            defaultContextMenu.MenuItems.Add(defaultMenuItem_1);

            //加入到应用程序级的右键菜单中
            Application.AddDefaultContextMenuExtension(defaultContextMenu);
        }
        
        [CommandMethod("McAddDefaultContextMenuItemUn")]
        public void McAddDefaultContextMenuItemUn()
        {
            if (defaultContextMenu != null)
            {
                if (defaultMenuItem_1 != null)
                {
                    defaultMenuItem_1.Click -= defaultMenuItem_1_Event;
                    defaultMenuItem_1 = null;
                }
                Application.RemoveDefaultContextMenuExtension(defaultContextMenu);
                defaultContextMenu = null;
            }
        }

        private static void MenuItem_2_Click(object sender, System.EventArgs e)
        {
            MyTopWindow.showModeless();
        }

        private static ContextMenuExtension objectContextMenu = null;
        private static MenuItem objectMenuItem_1 = null;
        private static EventHandler objectMenuItem_1_Event = new EventHandler(MenuItem_2_Click);

        //选定右键菜单
        [CommandMethod("McAddObjectContextMenuItem")]
        public void McAddObjectContextMenuItem()
        {
            //右键菜单 对象
            objectContextMenu = new ContextMenuExtension();
            objectContextMenu.Title = "test 选定后右键菜单";
            //右键菜单项及其事件
            objectMenuItem_1 = new MenuItem("Test 显示选定对象id");
            objectMenuItem_1.Click += objectMenuItem_1_Event;

            //菜单项添加到右键菜单
            objectContextMenu.MenuItems.Add(objectMenuItem_1);

            //加入到某一种对象的右键菜单中
            Application.AddObjectContextMenuExtension(Entity.GetClass(typeof(Entity)), objectContextMenu);
        }

        [CommandMethod("McAddObjectContextMenuItemUn")]
        public void McAddObjectContextMenuItemUn()
        {
            if (objectContextMenu != null)
            {
                if (objectMenuItem_1 != null)
                {
                    objectMenuItem_1.Click -= objectMenuItem_1_Event;
                    objectMenuItem_1 = null;
                }
                Application.RemoveObjectContextMenuExtension(Entity.GetClass(typeof(Entity)), objectContextMenu);
                objectContextMenu = null;
            }
        }

        //private static bool mainMenuClosed = false;
        //[CommandMethod("McCloseMainMenu")]
        //public void McCloseMainMenu()
        //{
        //    mainMenuClosed = true;
        //    //ToolPaletteManager.Manager.Schemes[0].UnloadCatalogs(CatalogTypeFlags.);

        //}

      
        [CommandMethod("McCustomMainMenuBizShow")]
        public void McCustomMainMenuBizShow()
        {
            //AcadApp.ActiveDocument.SendCommand("MENUBAR 1 ");

            string menuName = "自定义操作菜单";
            foreach (AcadPopupMenu acadPopupMenu in AcadApp.MenuGroups.Item(0).Menus)
            {
                if (menuName.Equals(acadPopupMenu.Name))
                {
                    if (!acadPopupMenu.OnMenuBar)
                        acadPopupMenu.InsertInMenuBar(AcadApp.MenuGroups.Item(0).Menus.Count);
                    return;
                }
            }

            // 创建菜单
            AcadPopupMenu customMainMenuBiz = customMainMenuBiz = AcadApp.MenuGroups.Item(0).Menus.Add(menuName);

            customMainMenuBiz.AddMenuItem(customMainMenuBiz.Count, "画线", "McNewLine ");
            customMainMenuBiz.AddMenuItem(customMainMenuBiz.Count, "画圆", "McNewCircle ");
            customMainMenuBiz.AddSeparator(customMainMenuBiz.Count);
            //CAD 界面切换
            customMainMenuBiz.AddMenuItem(customMainMenuBiz.Count, "CAD 界面切换", "McUiSwitch ");

            customMainMenuBiz.AddSeparator(customMainMenuBiz.Count);
            AcadPopupMenu subMenu = customMainMenuBiz.AddSubMenu(customMainMenuBiz.Count, "Test 子菜单");
            subMenu.AddMenuItem(customMainMenuBiz.Count, "Cricle", "circle ");
            customMainMenuBiz.AddSeparator(customMainMenuBiz.Count);
            customMainMenuBiz.AddMenuItem(customMainMenuBiz.Count, "Menu4", "rectangle ");

            customMainMenuBiz.InsertInMenuBar(AcadApp.MenuGroups.Item(0).Menus.Count);
        }

        [CommandMethod("McCustomMainMenuBizHide")]
        public void McCustomMainMenuBizHide()
        {
            string menuName = "自定义操作菜单";
            foreach (AcadPopupMenu acadPopupMenu in AcadApp.MenuGroups.Item(0).Menus)
            {
                if (menuName.Equals(acadPopupMenu.Name))
                {
                    if (!acadPopupMenu.OnMenuBar)
                        acadPopupMenu.InsertInMenuBar(AcadApp.MenuGroups.Item(0).Menus.Count);
                    return;
                }
            }
        }

        [CommandMethod("McCustomMainMenuSwitchShow")]
        public void McCustomMainMenuSwitchShow()
        {
            //AcadApp.ActiveDocument.SendCommand("MENUBAR 1 ");
            //Document doc = Application.DocumentManager.MdiActiveDocument;
            //doc.SendStringToExecute("MENUBAR 1 ", false,false, false);

            string menuName = "梦诚菜单";
            foreach (AcadPopupMenu acadPopupMenu in AcadApp.MenuGroups.Item(0).Menus)
            {
                if (menuName.Equals(acadPopupMenu.Name))
                {
                    if(!acadPopupMenu.OnMenuBar)
                        acadPopupMenu.InsertInMenuBar(AcadApp.MenuGroups.Item(0).Menus.Count);
                    return;
                }
            }

            // 创建菜单
            AcadPopupMenu customMainMenuSwitch = AcadApp.MenuGroups.Item(0).Menus.Add(menuName);
            //CAD 界面切换
            customMainMenuSwitch.AddMenuItem(0, "CAD 界面切换", "McUiSwitch ");
            customMainMenuSwitch.InsertInMenuBar(AcadApp.MenuGroups.Item(0).Menus.Count);
        }

        [CommandMethod("McCloseFileTabOnTop")]
        public void McCloseFileTabOnTop()
        {
            Commands.CloseFileTab();
        }

        [CommandMethod("McCloseFileTabOnTopUn")]
        public void McCloseFileTabOnTopUn()
        {
             Commands.DisplayFileTab();
        }

        [CommandMethod("McCloseRibbon")]
        public void McCloseRibbon()
        {
            if (RibbonServices.RibbonPaletteSet != null)
            {
                Commands.RibbonClose();
            }
        }

        [CommandMethod("McCloseRibbonUn")]
        public void McCloseRibbonUn()
        {
            Commands.Ribbon();
        }

        private static string cadWorkspaceName = null;
        
       /// <summary>
       /// 切换到 "AutoCAD 经典" 工作空间
       /// </summary>
       [CommandMethod("McWorkspace")]
       public void McWorkspace()
       {
           //string curCuiFileName = Application.GetSystemVariable("MENUNAME").ToString() + ".cuix";
           //CustomizationSection curCui = new CustomizationSection(curCuiFileName);
           string wsCurrentName = (string)Application.GetSystemVariable("WSCURRENT");

           if (!wsCurrentName.Equals(McWorkspaceName))
           {
               cadWorkspaceName = wsCurrentName;
                Application.SetCurrentWorkspace(McWorkspaceName);
            }

           //foreach (Workspace cuiWorkspace in curCui.Workspaces)
           //{
           //    //"WS_Anno2DDraft"  "草图与注释"
           //    //"WS_3DModeling   "三维建模"
           //    //"WS_SM_0001"     "三维基础"
           //    //"WS_AcadClassic" "AutoCAD 经典"
           //    if ("WS_AcadClassic".Equals(cuiWorkspace.ElementID) && !cuiWorkspace.Name.Equals(wsCurrentName))
           //    {
           //        Application.SetCurrentWorkspace(cuiWorkspace.Name);
           //        cadWorkspaceName = wsCurrentName;
           //        return;
           //    }
           //}
        }

        /// <summary>
        /// 切换到原工作空间
        /// </summary>
        [CommandMethod("McWorkspaceUn")]
        public void McWorkspaceUn()
        {
            if (cadWorkspaceName != null)
            {
                Application.SetCurrentWorkspace(cadWorkspaceName);
            }
        }

        private static void resetProfileInRegistry()
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey profilesKey = key.OpenSubKey(@"Software\Autodesk\AutoCAD\"+ REG_VERSION_SUB_KEY_ZHCN + @"\Profiles", true);
            if (profilesKey == null)
            {
                McWriteMessage("未安装 AutoCAD 2020 或未安装 AutoCAD2020 简体中文语言包，无法更新注册表默认配置文件设置");
                return;
            }

            Object val = profilesKey.GetValue("");
            if (val != null && PROFILE_MC2020.Equals(val.ToString(), StringComparison.CurrentCultureIgnoreCase))
            {
                string[] subNames = profilesKey.GetSubKeyNames();
                if (subNames != null && subNames.Length > 0)
                {
                    foreach (var subName in subNames)
                    {
                        if (!PROFILE_MC2020.Equals(subName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            profilesKey.SetValue("", subName);
                            return;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 使用新配置
        /// </summary>
        [CommandMethod("McSetActiveProfile")]
        public void McSetActiveProfile()
        {
            IAcadPreferences acadPreferences = (IAcadPreferences)Application.Preferences;

            if (!PROFILE_MC2020.Equals(acadPreferences.Profiles.ActiveProfile,
                StringComparison.CurrentCultureIgnoreCase)) //当前配置不是 MC2020 的话
            {
                object profileNames;
                acadPreferences.Profiles.GetAllProfileNames(out profileNames);
                IList profileNameList = (IList) profileNames;
                bool imported = false;
                foreach (var o in profileNameList)
                {
                    string s = o.ToString();
                    if (PROFILE_MC2020.Equals(s, StringComparison.CurrentCultureIgnoreCase))
                    {
                        imported = true;
                        break;
                    }
                }

                if (!imported)
                {
                    String dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    String profileRegFile = Path.Combine(dir, PROFILE_MC2020 + ".arg");
                    acadPreferences.Profiles.ImportProfile(PROFILE_MC2020, profileRegFile, true);
                }
                else
                {
                    acadPreferences.Profiles.ActiveProfile = PROFILE_MC2020;
                }

                
            }

            resetProfileInRegistry(); //注册表中记录的值重置为默认配置文件，AutoCAD 独立启动时仍可使用默认配置文件
        }

        [CommandMethod("McSetActiveProfileUn")]
        public void McSetActiveProfileUn()
        {
            IAcadPreferences acadPreferences = (IAcadPreferences)Application.Preferences;

            if (PROFILE_MC2020.Equals(acadPreferences.Profiles.ActiveProfile,
                StringComparison.CurrentCultureIgnoreCase))
            {
                object profileNames;
                acadPreferences.Profiles.GetAllProfileNames(out profileNames);
                IList profileNameList = (IList)profileNames;
                string foundS = null;
                foreach (var o in profileNameList)
                {
                    string s = o.ToString();
                    if (!PROFILE_MC2020.Equals(s, StringComparison.CurrentCultureIgnoreCase))
                    {
                        foundS  = s;
                        break;
                    }
                }

                if (foundS!=null)
                {
                    acadPreferences.Profiles.ActiveProfile = foundS;
                }

                //TODO
                //其他情况，需要逐个部分的处理
            }
        }

        [CommandMethod("McCustomCui")]
        public void McCustomCui()
        {
            bool loaded = false;

            for (int i = 0; i < AcadApp.MenuGroups.Count; i++)
            {

                if (MG_NAME_MCMENU.Equals(AcadApp.MenuGroups.Item(i).Name.ToUpper()))
                {
                    loaded = true;
                }
                else
                {
                    AcadApp.MenuGroups.Item(i).Unload();
                }
            }

            if (!loaded)
            {
                string dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                //保存的CUI文件名（从CAD2010开始，后缀改为了cuix）
                string strCuiFileName = dir + "\\" + MG_NAME_MCMENU + ".cuix";
                Application.LoadMainMenu(strCuiFileName);
            }
        }

        [CommandMethod("McCustomMainMenuClear")]
        public void McCustomMainMenuClear()

        {
            foreach (AcadPopupMenu acadPopupMenu in AcadApp.MenuGroups.Item(0).Menus)
            {
                if (acadPopupMenu.OnMenuBar)
                {
                    acadPopupMenu.RemoveFromMenuBar();
                }
            }
        }

        /// <summary>
        /// 代码生成 cuix 菜单配置
        /// </summary>
        [CommandMethod("SaveMenuToCuiX")]
        public void SaveMenuToCuiX()
        {
            ////自定义的组名
            string strMyGroupName = "McslGroup";
            String dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            //保存的CUI文件名（从CAD2010开始，后缀改为了cuix）
            string strCuiFileName = dir + "\\" + "" + ".cuix";

            //创建一个自定义组（这个组中将包含我们自定义的命令、菜单、工具栏、面板等）
            CustomizationSection myCSection = new CustomizationSection(strCuiFileName);
            myCSection.MenuGroupName = strMyGroupName;

            //创建自定义命令组
            MacroGroup mg = new MacroGroup("MCBIZ", myCSection.MenuGroup);
            MenuMacro mm1 = new MenuMacro(mg, "打开文件", "_MCOF", "");
            MenuMacro mm2 = new MenuMacro(mg, "打开模板", "_MCOM", "");
            MenuMacro mm3 = new MenuMacro(mg, "保存", "_MCSV", "");

            //声明菜单别名
            StringCollection scMyMenuAlias = new StringCollection();
            scMyMenuAlias.Add("MyPop1");
            scMyMenuAlias.Add("MyTestPop");

            //菜单项（将显示在项部菜单栏中）
            PopMenu pmParent = new PopMenu("MC菜单", scMyMenuAlias, "MC菜单", myCSection.MenuGroup);

            //子项的菜单（多级）
            PopMenu pm1 = new PopMenu("打开", new StringCollection(), "", myCSection.MenuGroup);
            PopMenuRef pmr1 = new PopMenuRef(pm1, pmParent, -1);
            PopMenuItem pmi1 = new PopMenuItem(mm1, "文件", pm1, -1);
            PopMenuItem pmi2 = new PopMenuItem(mm2, "模板", pm1, -1);

            //子项的菜单（单级）
            PopMenuItem pmi3 = new PopMenuItem(mm3, "保存(&S)", pmParent, -1);

            // 最后保存文件
            myCSection.Save();
        }
    }

}
