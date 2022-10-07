using System;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Platform.WindowManagement;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CloseTabsToRight.Helpers
{
    public static class WindowFrameHelpers
    {
        public static WindowFrame GetActiveWindowFrame(IEnumerable<IVsWindowFrame> frames, DTE2 dte)
        {
            return (from vsWindowFrame in frames
                    let window = GetWindow(vsWindowFrame)
                    where window == dte.ActiveWindow
                    select vsWindowFrame as WindowFrame)
                .FirstOrDefault();
        }

        public static Window GetWindow(IVsWindowFrame vsWindowFrame)
        {
            object window;
            ErrorHandler.ThrowOnFailure(vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject,
                out window));

            return window as Window;
        }

        public static IEnumerable<IVsWindowFrame> GetVsWindowFrames(IServiceProvider serviceProvider)
        {
            if(serviceProvider == null)
                throw new ArgumentNullException(nameof(serviceProvider));

            var windowFrames = new List<IVsWindowFrame>();

            if (!(serviceProvider.GetService(typeof(SVsUIShell)) is IVsUIShell uiShell))
            {
                return Enumerable.Empty<IVsWindowFrame>();
            }

            ErrorHandler.ThrowOnFailure(uiShell.GetDocumentWindowEnum(out var windowEnumerator));

            if (windowEnumerator.Reset() != VSConstants.S_OK)
                return Enumerable.Empty<IVsWindowFrame>();

            var frames = new IVsWindowFrame[1];
            var hasMorewindows = true;
            do
            {
                hasMorewindows = windowEnumerator.Next(1, frames, out var fetched) == VSConstants.S_OK && fetched == 1;

                if (!hasMorewindows || frames[0] == null)
                    continue;

                windowFrames.Add(frames[0]);

            } while (hasMorewindows);

            return windowFrames;
        }

        /// <summary>
        /// Some tabs, eg. a xaml tab, will have multiple window frames associated to it, with one of them being the root of other "sub-frames" in the heirarchy. 
        /// Sub-frames will have a null FrameView, so we need to return root as an active frame if the given active frame is a "sub-frame".
        /// </summary>
        /// <param name="allWindowFrames">Flat list of all window frames open in document group</param>
        /// <param name="activeFrame">Active frame that might be a "sub-frame"</param>
        /// <remarks>We need FrameView of a frame to be able to access DocumentGroup ancestor in heirarchy</remarks>
        public static WindowFrame GetRootFrameIfSubFrame(WindowFrame activeFrame, IEnumerable<WindowFrame> allWindowFrames)
        {
            if (activeFrame.FrameView == null && allWindowFrames.Contains(activeFrame.RootFrame))
            {
                return activeFrame.RootFrame;
            }

            return activeFrame;
        }
    }
}
