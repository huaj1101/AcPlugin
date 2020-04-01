// (C) Copyright 2020 by  
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

// This line is not mandatory, but improves loading performances
[assembly: ExtensionApplication(typeof(AutoCAD_CSharp_plug_in_acCustomUI.MyPlugin))]

namespace AutoCAD_CSharp_plug_in_acCustomUI
{

    // This class is instantiated by AutoCAD once and kept alive for the 
    // duration of the session. If you don't do any one time initialization 
    // then you should remove this class.
    public class MyPlugin : IExtensionApplication
    {

        void IExtensionApplication.Initialize()
        {
            // Add one time initialization here
            // One common scenario is to setup a callback function here that 
            // unmanaged code can call. 
            // To do this:
            // 1. Export a function from unmanaged code that takes a function
            //    pointer and stores the passed in value in a global variable.
            // 2. Call this exported function in this function passing delegate.
            // 3. When unmanaged code needs the services of this managed module
            //    you simply call acrxLoadApp() and by the time acrxLoadApp 
            //    returns  global function pointer is initialized to point to
            //    the C# delegate.
            // For more info see: 
            // http://msdn2.microsoft.com/en-US/library/5zwkzwf4(VS.80).aspx
            // http://msdn2.microsoft.com/en-us/library/44ey4b32(VS.80).aspx
            // http://msdn2.microsoft.com/en-US/library/7esfatk4.aspx
            // as well as some of the existing AutoCAD managed apps.

            // Initialize your plug-in application here

            ////MessageBox.Show("插件已加载");
            ///
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("MyPlugin插件初始化\n");

            //TODO 处理工作区的 tab - 应不可关闭

            //TODO 处理左上角图标和主菜单
            //MyCommands.ZsyMainMenu();

            
            //无选定对象时的右键菜单
            MyCommands.ZsyAddDefaultContextMenuItem();

            editor.WriteMessage("1 完成 MyCommands.ZsyAddDefaultContextMenuItem()\n");

            //有选定对象时的右键菜单
            MyCommands.ZsyAddObjectContextMenuItem();
            editor.WriteMessage("2  MyCommands.ZsyAddObjectContextMenuItem()\n");
            
            //顶部Ribbon 菜单区域，AutoCAD自带tab隐藏，添加自定义tab、按钮
            MyCommands.ZsyRibbonTab();
            editor.WriteMessage("3  MyCommands.ZsyRibbonTab()\n");
            //

            //上下左右添加自定义停靠区域
            MyCommands.ZsyPaletteSetLeft();  //包含弹出模态窗口、非模态置顶窗口
            editor.WriteMessage("4  MyCommands.ZsyPaletteSetLeft()\n");
            MyCommands.ZsyPaletteSetTop();
            editor.WriteMessage("5  MyCommands.ZsyPaletteSetTop()\n");
            MyCommands.ZsyPaletteSetBottom();
            editor.WriteMessage("6  MyCommands.ZsyPaletteSetBottom()\n");
            MyCommands.ZsyPaletteSetRight();  //WPF UserControl
            editor.WriteMessage("7  MyCommands.ZsyPaletteSetRight()\n");

            //监听选定事件，选定对象id 在 非模态置顶窗口中显示
            MyEvent.AddDocEvent();
            editor.WriteMessage("8  MyEvent.AddDocEvent()\n");


            //隐藏AutoCAD自带 toolbar，添加自定义toolbar
            MyCommands.ZsyToolbar();
            editor.WriteMessage("9  MyCommands.ZsyToolbar()\n");
        }

        void IExtensionApplication.Terminate()
        {
            // Do plug-in application clean up here
        }

    }

}
