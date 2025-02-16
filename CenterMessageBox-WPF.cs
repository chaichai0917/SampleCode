﻿//------------------------------------------------------------------------------
// MessageBox - 親画面中央表示とXボタン非表示（WPF）
//------------------------------------------------------------------------------
// [NOTE] 
//   親画面中央表示は、下記情報を利用
//   https://millyc.hatenadiary.org/entry/20080312/1205311545
//
//   Xボタン非表示
//   Windows Server 系で、Xボタンディセーブルが視認できないため
//
//   Windows Forms → WPF
//
//------------------------------------------------------------------------------
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;


namespace MyTools
{
    public class CenterMessageBox
    {
        #region defines

        private static class NativeMethods
        {
            // Xボタン非表示
            [DllImport("user32.dll")]
            public static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, long dwLong);

            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);

            [DllImport("kernel32.dll")]
            public static extern IntPtr GetCurrentThreadId();

            [DllImport("user32.dll")]
            public static extern IntPtr SetWindowsHookEx(int idHook, HOOKPROC lpfn, IntPtr hInstance, IntPtr threadId);

            [DllImport("user32.dll")]
            public static extern bool UnhookWindowsHookEx(IntPtr hHook);

            [DllImport("user32.dll")]
            public static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

            [DllImport("user32.dll")]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

            public delegate IntPtr HOOKPROC(int nCode, IntPtr wParam, IntPtr lParam);

            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            public struct RECT
            {
                #region fields

                public int Left;
                public int Top;
                public int Right;
                public int Bottom;

                #endregion

                #region properties

                public int Width
                {
                    get { return this.Right - this.Left; }
                }

                public int Height
                {
                    get { return this.Bottom - this.Top; }
                }

                #endregion

                #region methods

                public Rectangle ToRectangle()
                {
                    return new Rectangle(this.Left, this.Top, this.Width, this.Height);
                }

                #endregion
            }

            public const int GWL_HINSTANCE = (-6);
            public const int WH_CBT = 5;
            public const int HCBT_ACTIVATE = 5;
            public const int SWP_NOSIZE = 0x0001;
            public const int SWP_NOZORDER = 0x0004;
            public const int SWP_NOACTIVATE = 0x0010;
            public const int GWL_STYLE = (-16);        // Xボタン非表示
            public const int WS_SYSMENU = 0x00080000;  // Xボタン非表示
        }

        #endregion

        #region static methods

        public static MessageBoxResult Show(Window owner, string text)
        {
            return Show(owner, text, string.Empty, MessageBoxButton.OK, MessageBoxImage.None, MessageBoxResult.None);
        }

        public static MessageBoxResult Show(Window owner, string text, string caption)
        {
            return Show(owner, text, caption, MessageBoxButton.OK, MessageBoxImage.None, MessageBoxResult.None);
        }

        public static MessageBoxResult Show(Window owner, string text, string caption, MessageBoxButton buttons)
        {
            return Show(owner, text, caption, buttons, MessageBoxImage.None, MessageBoxResult.None);
        }

        public static MessageBoxResult Show(Window owner, string text, string caption, MessageBoxButton buttons, MessageBoxImage icon)
        {
            return Show(owner, text, caption, buttons, icon, MessageBoxResult.None);
        }

        public static MessageBoxResult Show(Window owner, string text, string caption, MessageBoxButton buttons, MessageBoxImage icon, MessageBoxResult defaultResult)
        {
            if (null == owner)
            {
                throw new ArgumentNullException("owner");
            }
            CenterMessageBox messageBox = new CenterMessageBox(owner);
            return messageBox.Show(text, caption, buttons, icon, defaultResult);
        }

        #endregion

        #region fields

        private readonly Window Owner;
        private IntPtr HookHandle = IntPtr.Zero;
        private MessageBoxButton HookButtons;     // Xボタン非表示

        #endregion

        #region constructors

        private CenterMessageBox(Window owner)
        {
            this.Owner = owner;
        }

        #endregion

        #region methods

        private MessageBoxResult Show(
            string text,
            string caption,
            MessageBoxButton buttons,
            MessageBoxImage icon,
            MessageBoxResult defaultResult)
        {
            HwndSource hwndSource = (HwndSource)HwndSource.FromVisual(Owner);
            IntPtr hInstance = NativeMethods.GetWindowLong(hwndSource.Handle, NativeMethods.GWL_HINSTANCE);
            IntPtr threadId = NativeMethods.GetCurrentThreadId();
            this.HookHandle = NativeMethods.SetWindowsHookEx(NativeMethods.WH_CBT, this.HookProc, hInstance, threadId);
            this.HookButtons = buttons;  // Xボタン無効化

            return MessageBox.Show(this.Owner, text, caption, buttons, icon, defaultResult);
        }

        private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == NativeMethods.HCBT_ACTIVATE)
            {
                NativeMethods.RECT ownerRect;
                NativeMethods.RECT msgBoxRect;

                HwndSource hwndSource = (HwndSource)HwndSource.FromVisual(Owner);
                NativeMethods.GetWindowRect(hwndSource.Handle, out ownerRect);
                NativeMethods.GetWindowRect(wParam, out msgBoxRect);
                int x = ownerRect.Left + (ownerRect.Width - msgBoxRect.Width) / 2;
                int y = ownerRect.Top + (ownerRect.Height - msgBoxRect.Height) / 2;

                Rect workingArea = SystemParameters.WorkArea;
                if (workingArea.Bottom < y + msgBoxRect.Height)
                {
                    y = (int)workingArea.Bottom - msgBoxRect.Height;
                }
                if (workingArea.Right < x + msgBoxRect.Width)
                {
                    x = (int)workingArea.Right - msgBoxRect.Width;
                }
                if (y < workingArea.Top)
                {
                    y = (int)workingArea.Top;
                }
                if (x < workingArea.Left)
                {
                    x = (int)workingArea.Left;
                }

                NativeMethods.SetWindowPos(wParam, 0, x, y, 0, 0,
                    NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);

                // Xボタン非表示
                if (this.HookButtons == MessageBoxButton.YesNo)
                {
                    long style = (long)NativeMethods.GetWindowLong(wParam, NativeMethods.GWL_STYLE);
                    NativeMethods.SetWindowLong(wParam, NativeMethods.GWL_STYLE, style & ~NativeMethods.WS_SYSMENU);
                }

                try
                {
                    return NativeMethods.CallNextHookEx(this.HookHandle, nCode, wParam, lParam);
                }
                finally
                {
                    NativeMethods.UnhookWindowsHookEx(this.HookHandle);
                    this.HookHandle = IntPtr.Zero;
                }
            }

            return NativeMethods.CallNextHookEx(this.HookHandle, nCode, wParam, lParam);
        }

        #endregion
    }
}

