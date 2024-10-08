﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using WebSocketSharp.Server;
using WebSocketSharp;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Data;
using System.IO;
using Microsoft.Office.Core;
using System.Diagnostics;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointWSRemote
{
    public partial class PowerPointWSAddIn
    {
        public static PowerPointWSAddIn instance;
        private WebSocketServer server;

        private bool slideShowClosedFlag = false;
        private void PowerPointWSAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
            this.Setup();


            Application.SlideShowNextSlide += OnNextSlide;
            Application.SlideShowBegin += OnSlideShowBegin;
            Application.SlideShowEnd += OnSlideShowEnd;
        }

        private void PowerPointWSAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (server != null && server.IsListening)
            {
                server.Stop();
            }
        }

        public void Setup()
        {
            if (server != null)
            {
                server.Stop();
                server = null;
            }
            if (Properties.Settings.Default.enabled == false) return; 
            int port = int.Parse(Properties.Settings.Default.port);
            server = new WebSocketServer(port);
            server.AddWebSocketService<WebSocketHandler>("/ws");
            server.Start();
        }

        private void OnNextSlide(SlideShowWindow Wn)
        {
            this.SendStatus();
        }

        private void OnSlideShowBegin(SlideShowWindow wn)
        {
            this.SendStatus();
        }

        private void OnSlideShowEnd(Presentation presentation)
        {
            this.slideShowClosedFlag = true;
            this.SendStatus();
        }

        public enum MediaControlFunction
        {
            PLAY,
            PAUSE,
            STOP
        }



        public List<Shape> GetCurrentSlideMediaShapes()
        {
            if (Application.Presentations.Count == 0) return null;
            if (Application.ActivePresentation == null) return null;


            List<Shape> mediaShapes = new List<Shape>();

            try
            {
                foreach (Shape shape in this.Application.ActivePresentation.SlideShowWindow.View.Slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoMedia)
                    {
                        if (shape.MediaType == PpMediaType.ppMediaTypeMovie || shape.MediaType == PpMediaType.ppMediaTypeSound)
                        {
                            mediaShapes.Add(shape);
                        }
                    }
                }
                if (mediaShapes.Count > 0)
                {
                    return mediaShapes;
                }
            }
            catch (Exception e)
            {
                return null;
            }


            return null;
        }

       

        

        public void BeginPresentation()
        {
            if (Application.Presentations.Count == 0) return;
            if (Application.ActivePresentation == null) return;

            Presentation presentation = Application.ActivePresentation;
            presentation.SlideShowSettings.Run();
        }

        public void EndPresentation()
        {
            ;
            if (Application.Presentations.Count == 0) return;
            if (Application.SlideShowWindows.Count == 0) return;
            Application.SlideShowWindows[1].View.Exit();
        }



        public void OpenPresentation(string path, bool closeOthers)
        {
            string fullPath = Path.GetFullPath(path);

            Presentation existingPresentation = null;

            // Check if the presentation is already open
            foreach (Presentation pres in Application.Presentations)
            {
                if (string.Equals(pres.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                {
                    existingPresentation = pres;
                    break;
                }
            }

            // If the presentation is already open
            if (existingPresentation != null)
            {
                if (closeOthers)
                {
                    foreach (Presentation pres in Application.Presentations)
                    {
                        // Close all presentations except the already open one
                        if (pres != existingPresentation)
                        {
                            pres.Close();
                        }
                    }
                }

                // Activate the window with the already open presentation
                foreach (DocumentWindow window in Application.Windows)
                {
                    if (window.Presentation == existingPresentation)
                    {
                        window.Activate();
                        return;
                    }
                }

                // If no window was found for the presentation, activate the first window (fallback)
                if (Application.Windows.Count > 0)
                {
                    Application.Windows[1].Activate();
                }
                return;
            }

            // If the presentation was not found, open it
            if (closeOthers)
            {
                this.CloseAllPresentations();
            }

            try
            {
                Application.Presentations.Open(fullPath);
            }
            catch (Exception)
            {

            }
        }



        public void CloseAllPresentations()
        {
            var presentations = Application.Presentations;

            var presentationsToClose = new List<Presentation>();

            foreach (Presentation presentation in presentations)
            {
                presentationsToClose.Add(presentation);
            }

            foreach (Presentation presentation in presentationsToClose)
            {
                try
                {
                    presentation.Close();
                }
                catch (Exception)
                {
                }
            }
        }

        public void NextSlide()
        {
            try
            {
                Application.ActivePresentation.SlideShowWindow.View.Next();
            }
            catch (COMException)
            {

            }
        }

        public void PrevSlide()
        {
            try
            {
                Application.ActivePresentation.SlideShowWindow.View.Previous();
            }
            catch (COMException)
            {

            }
        }

        public void GoToSlide(int slideNumber)
        {
            try
            {
                Application.ActivePresentation.SlideShowWindow.View.GotoSlide(slideNumber);

            }
            catch (COMException)
            {

            }
        }

        public int GetCurrentSlide()
        {
            int result = 0;
            try
            {
                result = this.Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition;
            }
            catch (COMException)
            {

            }

            return result;
        }
        
        public int GetTotalSlideCount()
        {
            int result = 0;
            try
            {
                result = this.Application.ActivePresentation.Slides.Count;
            }
            catch (COMException)
            {

            }

            return result;
        }

        public void EraseDrawings()
        {
            try
            {
                Application.ActivePresentation.SlideShowWindow.View.EraseDrawing();

            }
            catch (COMException)
            {

            }
        }

        

        public void ToggleLaserPointer(MsoTriState option)
        {
            try
            {
                if(option == MsoTriState.msoTriStateToggle)
                {
                    if (((dynamic)Application.ActivePresentation.SlideShowWindow.View).LaserPointerEnabled == true)
                        
                    {
                        ((dynamic)Application.ActivePresentation.SlideShowWindow.View).LaserPointerEnabled = false;
                    }
                    else
                    {
                        ((dynamic)Application.ActivePresentation.SlideShowWindow.View).LaserPointerEnabled = true;
                    }
                    return;
                }
                ((dynamic)Application.ActivePresentation.SlideShowWindow.View).LaserPointerEnabled = option;
            }
            catch (COMException)
            {

            }
        }

        public void BlackOutPresentation(PpSlideShowState blackoutOption)
        {
            try
            {
                Application.ActivePresentation.SlideShowWindow.View.State = blackoutOption;
            }
            catch (COMException)
            {

            }
        }

        public void HideSlide(int slideId)
        {
            try
            {
                Application.ActivePresentation.Slides[slideId].SlideShowTransition.Hidden = MsoTriState.msoTrue;
            }
            catch (COMException)
            {
              
            }
        }
        public void UnhideSlide(int slideId)
        {
            try
            {
                Application.ActivePresentation.Slides[slideId].SlideShowTransition.Hidden = MsoTriState.msoFalse;
            }
            catch (COMException)
            {

            }
        }

        public void UnhideAllSlides()
        {
            try
            {
                // Ensure there's an active presentation
                if (this.Application.ActivePresentation != null)
                {
                    // Get the collection of slides in the active presentation
                    Slides slides = this.Application.ActivePresentation.Slides;

                    // Iterate through all slides
                    foreach (Slide slide in slides)
                    {
                        // Unhide the slide
                        slide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
                    }
                }
            }
            catch (COMException)
            {

            }
        }


        public void SendStatus()
        {
            JObject response = new JObject();

            response["slideShowActive"] = false;
            try
            {
                if (Application.ActivePresentation.SlideShowWindow != null)
                {
                    response["slideShowActive"] = true;
                }
            }
            catch (Exception e)
            {

            }
            if (this.slideShowClosedFlag)
            {
                response["slideShowActive"] = false;
                this.slideShowClosedFlag = false;
            }
            response["totalSlideCount"] = this.GetTotalSlideCount();
            response["currentSlide"] = this.GetCurrentSlide();
            List<Shape> currentSlideShapes = this.GetCurrentSlideMediaShapes();
            if (currentSlideShapes != null) response["currentSlideMediaCount"] = currentSlideShapes.Count;
            else response["currentSlideMediaCount"] = 0;
            response["slideNotes"] = null;
            try
            {
                Slide slide = this.Application.ActivePresentation.SlideShowWindow.View.Slide;
                if (slide.HasNotesPage == MsoTriState.msoTrue)
                {
                    SlideRange notesPages = slide.NotesPage;
                    foreach (PowerPoint.Shape shape in notesPages.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoPlaceholder)
                        {
                            if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                            {
                                response["slideNotes"] = shape.TextFrame.TextRange.Text;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }
            response["presentationFullPath"] = null;
            response["fileName"] = null;
            try
            {
                response["presentationFullPath"] = Application.ActivePresentation.FullName;
                response["fileName"] = Application.ActivePresentation.Name;
            }
            catch (Exception e)
            {

            }

            SendWebSocketMessage(response.ToString());
        }


        private void SendWebSocketMessage(string message)
        {
            if (server != null && server.IsListening)
            {
                var webSocketBehavior = server.WebSocketServices["/ws"];

                var activeSessions = webSocketBehavior.Sessions.Sessions.ToList();

                foreach (var session in activeSessions)
                {
                    session.Context.WebSocket.Send(message);
                }
            }
        }

        private class WebSocketHandler : WebSocketBehavior
        {
            protected override void OnOpen()
            {
                PowerPointWSAddIn.instance.SendStatus();
            }
            protected override void OnMessage(MessageEventArgs e)
            {
                string receivedMessage = e.Data;

                try
                {
                    JObject jsonObject = JObject.Parse(receivedMessage);

                    if (jsonObject["slideShowActive"] != null)
                    {
                        bool slideShowActiveState = bool.Parse(jsonObject["slideShowActive"].ToString());
                        if (slideShowActiveState == true)
                        {
                            PowerPointWSAddIn.instance.BeginPresentation();
                        }
                        else if (slideShowActiveState == false)
                        {
                            PowerPointWSAddIn.instance.EndPresentation();
                        }
                    }


                    if (jsonObject["currentSlide"] != null)
                    {
                        int slideNumber = int.Parse(jsonObject["currentSlide"].ToString());
                        PowerPointWSAddIn.instance.GoToSlide(slideNumber);
                    }

                    if (jsonObject["action"] != null)
                    {
                        string action = jsonObject["action"].ToString();
                        if (action == "next")
                        {
                            PowerPointWSAddIn.instance.NextSlide();
                        }
                        else if (action == "previous")
                        {
                            PowerPointWSAddIn.instance.PrevSlide();
                        }
                        else if (action == "first")
                        {
                            PowerPointWSAddIn.instance.GoToSlide(1);
                        }
                        else if (action == "last")
                        {
                            try
                            {
                                PowerPointWSAddIn.instance.GoToSlide(PowerPointWSAddIn.instance.Application.ActivePresentation.Slides.Count);
                            }
                            catch (COMException ex)
                            {

                            }
                        }
                        else if (action == "status")
                        {
                            PowerPointWSAddIn.instance.SendStatus();
                        }
                        else if (action == "closeAll")
                        {
                            PowerPointWSAddIn.instance.CloseAllPresentations();
                            PowerPointWSAddIn.instance.SendStatus();
                        }
                        else if (action == "openPresentation")
                        {
                            if (jsonObject["path"] != null)
                            {
                                bool closeOthers = true;
                                if (jsonObject["closeOthers"] != null)
                                {
                                    closeOthers = bool.Parse(jsonObject["closeOthers"].ToString());
                                }
                                string path = jsonObject["path"].ToString();
                                PowerPointWSAddIn.instance.OpenPresentation(path, closeOthers);
                             
                            }
                            PowerPointWSAddIn.instance.SendStatus();
                        }
                        
                        else if (action == "blackout")
                        {
                            PowerPointWSAddIn.instance.BlackOutPresentation(PpSlideShowState.ppSlideShowBlackScreen);
                        }
                        else if (action == "whiteout")
                        {
                            PowerPointWSAddIn.instance.BlackOutPresentation(PpSlideShowState.ppSlideShowWhiteScreen);
                        }
                        else if (action == "showPresentation")
                        {
                            PowerPointWSAddIn.instance.BlackOutPresentation(PpSlideShowState.ppSlideShowRunning);
                        }
                        else if (action == "hideSlide")
                        {
                            if (jsonObject["slideId"] != null)
                            {
                                int slideId;
                                if (int.TryParse(jsonObject["slideId"].ToString(), out slideId) && slideId >= 0)
                                {
                                    PowerPointWSAddIn.instance.HideSlide(slideId);
                                }
                            }
                        }
                        else if (action == "unhideSlide")
                        {
                            if (jsonObject["slideId"] != null)
                            {
                                int slideId;
                                if (int.TryParse(jsonObject["slideId"].ToString(), out slideId) && slideId >= 0)
                                {
                                    PowerPointWSAddIn.instance.UnhideSlide(slideId);
                                }
                            }
                        }
                        else if (action == "unhideAllSlides")
                        {
                            PowerPointWSAddIn.instance.UnhideAllSlides();
                        }
                        else if (action == "showLaserPointer")
                        {
                            PowerPointWSAddIn.instance.ToggleLaserPointer(MsoTriState.msoTrue);
                        }
                        else if (action == "hideLaserPointer")
                        {
                            PowerPointWSAddIn.instance.ToggleLaserPointer(MsoTriState.msoFalse);
                        }
                        else if (action == "toggleLaserPointer")
                        {
                            PowerPointWSAddIn.instance.ToggleLaserPointer(MsoTriState.msoTriStateToggle);
                        }
                        else if (action == "eraseDrawings")
                        {
                            PowerPointWSAddIn.instance.EraseDrawings();
                        }
                    }
                }
                catch (JsonReaderException ex)
                {
                    Send("Error: " + ex.Message);
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(PowerPointWSAddIn_Startup);
            this.Shutdown += new System.EventHandler(PowerPointWSAddIn_Shutdown);
        }

        #endregion
    }
}
