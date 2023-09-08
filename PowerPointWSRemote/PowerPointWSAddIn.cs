using System;
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
            Application.WindowSelectionChange += OnWindowChanged;
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
            if(server != null)
            {
                server.Stop();
                server = null;
            }
            int port = int.Parse(Properties.Settings.Default.port);
            server = new WebSocketServer(port);
            server.AddWebSocketService<WebSocketHandler>("/ws");
            server.Start();
        }

        private void OnWindowChanged(Selection Sel)
        {
            this.SendStatus();
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

        public void BeginPresentation()
        {
            if (Application.Presentations.Count == 0) return;
            if (Application.ActivePresentation == null) return;

            Presentation presentation = Application.ActivePresentation;
            presentation.SlideShowSettings.Run();
        }

        public void EndPresentation()
        {
            if (Application.Presentations.Count == 0) return;
            if (Application.SlideShowWindows.Count == 0) return;
            Application.SlideShowWindows[1].View.Exit();
        }

        public void OpenPresentation(string path, bool closeOthers)
        {
            if(closeOthers)
            {
                this.CloseAllPresentations();
            }
            
            try
            {
                Application.Presentations.Open(@path);

            }
            catch(FileNotFoundException e)
            {
                this.SendWebSocketMessage("File not found");
            }
        }

        public void CloseAllPresentations()
        {
            foreach (Presentation pres in Application.Presentations)
            {
                pres.Close();
            }
            try
            {
                Application.ActivePresentation.Close();
            }
            catch (Exception e)
            {

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
            if(this.slideShowClosedFlag)
            {
                response["slideShowActive"] = false;
                this.slideShowClosedFlag = false;
            }
            response["totalSlideCount"] = this.GetTotalSlideCount();
            response["currentSlide"] = this.GetCurrentSlide();
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

            


            string jsonResponse = response.ToString();

            SendWebSocketMessage(jsonResponse);
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

                    if(jsonObject["slideShowActive"] != null)
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
                        if(action == "next")
                        {
                            PowerPointWSAddIn.instance.NextSlide();
                        }
                        else if(action == "previous")
                        {
                            PowerPointWSAddIn.instance.PrevSlide();
                        }
                        else if(action == "first")
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
