using Google.Apis.Auth.OAuth2;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Drive.v3;

namespace SlidesDemo
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/slides.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SlidesService.Scope.Presentations,DriveService.Scope.Drive };
        static string ApplicationName = "Google Slides API .NET Quickstart";
       // static string presentationId;

        static void Main(string[] args)
        {
            string presentationName = "MyPresentation2";
            String slideName = "MyNewSlide_001";
            String imagetextBoxId = "MyTextBox_02";
            string textTextBoxId = "MyTextBox_01";
            string videotextTextBoxId = "MyTextBox_03";

            CreateBlankPresentation(presentationName);

            var presentationId =  GetPresentationByName(presentationName);
            SetTitleAndSubtitle(presentationId, "Title", "Sub title");

            var slideId = GetFirstSlideId(presentationId);
            AddVideoToSlide(presentationId,slideId, videotextTextBoxId);
            // var slideId2 = AddSlideToPresentation(presentationId, slideName);
            AddSlideBackground(presentationId, slideId);

           // AddTextToSlide(presentationId,slideId, textTextBoxId);

         
            //AddImageToSlide(presentationId, slideId, imagetextBoxId);

           // AddBackgroundImageToSlide(presentationId, slideId);

            DisplayPresentation();
        }

        private static void SetTitleAndSubtitle(string presentationId, string title, string subtitle)
        {
            SlidesService slidesService = GetSlideServiceClient();

            var slide = GetSlides(presentationId).FirstOrDefault();
            var titleObject = slide.PageElements.FirstOrDefault(s => s.Shape.Placeholder.Type == "CENTERED_TITLE");
            var subtitleObject = slide.PageElements.FirstOrDefault(s => s.Shape.Placeholder.Type == "SUBTITLE");
            List<Request> requests = new List<Request>();
            Request request = new Request();

            requests.Add(ChangeText(title, titleObject));
            requests.Add(ChangeText(subtitle, subtitleObject));

            //requests.Add(request);

            // If you wish to populate the slide with elements, add create requests here,
            // using the slide ID specified above.

            // Execute the request.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate(body, presentationId).Execute();
            //CreateSlideResponse createSlideResponse = response.Replies.First().CreateSlide;
            //Console.WriteLine("Created slide with ID: " + createSlideResponse.ObjectId);
            //return createSlideResponse.ObjectId;

        }

        private static Request ChangeText(string title, PageElement titleObject )
        {
            Request request = new Request();
            //var slideRequest = new DeleteTextRequest();
            //slideRequest.ObjectId = titleObject.ObjectId;
            //slideRequest.TextRange = new Range() { Type = "All" };
            //request.DeleteText = (slideRequest);

            request.InsertText = (new InsertTextRequest()
            {
                ObjectId = titleObject.ObjectId,
                InsertionIndex = (0),
                Text = title
            });
            //request.ReplaceAllText = new ReplaceAllTextRequest()
            //{
            //    ContainsText = titleObject.
            //}
            return request;
        }

        private static void AddVideoToSlide(string presentationId, string slideId, string videotextTextBoxId)
        {
            SlidesService slidesService = GetSlideServiceClient();

            var slide = GetSlides(presentationId).FirstOrDefault(s => s.ObjectId == slideId);
            // Create a new square text box, using a supplied object ID.
            List<Request> requests = new List<Request>();
            Dimension pt350 = new Dimension() { Magnitude = 350.0, Unit = "PT" };
            var afflineTransform = new AffineTransform()
            {
                ScaleX = (1.0),
                ScaleY = (1.0),
                TranslateX = (350.0),
                TranslateY = (100.0),
                Unit = ("PT")
            };
            Size size = slide.PageElements.FirstOrDefault().Size;
            var pageElementProperties = new PageElementProperties()
            {
                PageObjectId = slideId,
                Size = size,
                Transform = afflineTransform
            };
          
            // Insert text into the box, using the object ID given to it.
            requests.Add(new Request()
            {
                
                CreateVideo = new CreateVideoRequest()
                {
                    ObjectId = videotextTextBoxId,
                    Id = "z3iMVQlcuoE",
                    Source = "YOUTUBE", 
                    ElementProperties = pageElementProperties
                }
            });

            // Execute the requests.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate(body, presentationId).Execute();
            CreateVideoResponse createShapeResponse = response.Replies.First().CreateVideo;
            Console.WriteLine("Created video with ID: " + createShapeResponse.ObjectId);
        }

        private static void AddImageToSlide(string presentationId, string slideId,string textBoxId)
        {
           SlidesService slidesService = GetSlideServiceClient();
            // Create a new square text box, using a supplied object ID.
            List<Request> requests = new List<Request>();
            Dimension pt350 = new Dimension() { Magnitude = 350.0, Unit = "PT" };
            var afflineTransform = new AffineTransform()
            {
                ScaleX = (1.0),
                ScaleY = (1.0),
                TranslateX = (350.0),
                TranslateY = (100.0),
                Unit = ("PT")
            };
            Size size = new Size()
            {
                Height = (pt350),
                Width = (pt350)
            };
            var pageElementProperties= new PageElementProperties()
            {
                PageObjectId = slideId,
                Size = size,
                Transform = afflineTransform
            };
            //requests.Add(new Request()
            //{
            //    CreateShape = new CreateShapeRequest()
            //    {
            //        ObjectId = textBoxId,
            //        ShapeType = "TEXT_BOX",
            //        ElementProperties = pageElementProperties
            //    }
            //});

            // Insert text into the box, using the object ID given to it.
            requests.Add(new Request()
            {
                CreateImage = new CreateImageRequest()
                {
                    ObjectId = textBoxId,
                    Url = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4gPTyRGSqYQoObgXWS9JRdsqBI7sQfJfgaAwQiZR43GxrR2hjWJE-Hg",
                    ElementProperties = pageElementProperties
                }
            });

            // Execute the requests.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate( body, presentationId).Execute();
            CreateImageResponse createShapeResponse = response.Replies.First().CreateImage;
            Console.WriteLine("Created textbox with ID: " + createShapeResponse.ObjectId);
        }

        private static string GetPresentationByName(string presentationName)
        {
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GetCredentials(),
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;

            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                    if (file.Name == presentationName)
                        return file.Id;
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            return null;

        }

        private static void AddTextToSlide(string presentationId, string slideId, String textBoxId)
        {
            SlidesService slidesService = GetSlideServiceClient();
            // Create a new square text box, using a supplied object ID.
            List<Request> requests = new List<Request>();
            Dimension pt350 = new Dimension() { Magnitude = 350.0, Unit = "PT" };
            var afflineTransform = new AffineTransform()
            {
                ScaleX = (1.0),
                ScaleY = (1.0),
                TranslateX = (350.0),
                TranslateY = (100.0),
                Unit = ("PT")
            };
            Size size = new Size()
            {
                Height = (pt350),
                Width = (pt350)
            };
            var pageElementProperties= new PageElementProperties()
            {
                PageObjectId = slideId,
                Size = size,
                Transform = afflineTransform
            };
            requests.Add(new Request()
            {
                CreateShape = new CreateShapeRequest()
                {
                    ObjectId = (textBoxId),
                    ShapeType = ("TEXT_BOX"),
                    ElementProperties = pageElementProperties
                }
            });

            // Insert text into the box, using the object ID given to it.
            requests.Add(new Request()
            {
                InsertText = (new InsertTextRequest()
                {
                    ObjectId = (textBoxId),
                    InsertionIndex = (0),
                    Text = ("New Box Text Inserted")
                })
            });

            // Execute the requests.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate( body, presentationId).Execute();
            CreateShapeResponse createShapeResponse = response.Replies.First().CreateShape;
            Console.WriteLine("Created textbox with ID: " + createShapeResponse.ObjectId);

        }

        private static string AddSlideToPresentation(string presentationId,string slideId)
        {
            SlidesService slidesService = GetSlideServiceClient();
            // Add a slide at index 1 using the predefined "TITLE_AND_TWO_COLUMNS" layout
            // and the ID "MyNewSlide_001".
            List<Request> requests = new List<Request>();
            Request request = new Request();
            var slideRequest = new CreateSlideRequest();
            slideRequest.ObjectId = (slideId);
            slideRequest.InsertionIndex = (1);
            slideRequest.SlideLayoutReference = (new LayoutReference() { PredefinedLayout = "TITLE_AND_TWO_COLUMNS" });
           
            request.CreateSlide = (slideRequest);
            requests.Add(request);

            // If you wish to populate the slide with elements, add create requests here,
            // using the slide ID specified above.

            // Execute the request.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate( body, presentationId).Execute();
            CreateSlideResponse createSlideResponse = response.Replies.First().CreateSlide;
            Console.WriteLine("Created slide with ID: " + createSlideResponse.ObjectId);
            return createSlideResponse.ObjectId;
        }
        private static void AddSlideBackground(string presentationId, string slideId)
        {
            SlidesService slidesService = GetSlideServiceClient();
            // Add a slide at index 1 using the predefined "TITLE_AND_TWO_COLUMNS" layout
            // and the ID "MyNewSlide_001".
            List<Request> requests = new List<Request>();
            Request request = new Request();
            
            var slideRequest = new UpdatePagePropertiesRequest();
            slideRequest.ObjectId = (slideId);
            slideRequest.Fields = "pageBackgroundFill";
            slideRequest.PageProperties = new PageProperties()
            {
                PageBackgroundFill = new PageBackgroundFill()
                {
                    StretchedPictureFill = new StretchedPictureFill()
                    {
                        ContentUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4gPTyRGSqYQoObgXWS9JRdsqBI7sQfJfgaAwQiZR43GxrR2hjWJE-Hg"
                    }
                }
            };
            request.UpdatePageProperties = slideRequest;
            //slideRequest.InsertionIndex = (1);
            //slideRequest.SlideLayoutReference = (new LayoutReference() { PredefinedLayout = "TITLE_AND_TWO_COLUMNS" });
            //request.UpdateShapeProperties = new UpdateShapePropertiesRequest()
        //    slideRequest.UpdatePageProperties = new UpdatePagePropertiesRequest() {Fields= "pageBackgroundFill" , PageProperties = new PageProperties() { PageBackgroundFill = new PageBackgroundFill() { StretchedPictureFill = new StretchedPictureFill() { ContentUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4gPTyRGSqYQoObgXWS9JRdsqBI7sQfJfgaAwQiZR43GxrR2hjWJE-Hg" } } } };
            requests.Add(request);

            // If you wish to populate the slide with elements, add create requests here,
            // using the slide ID specified above.

            // Execute the request.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate(body, presentationId).Execute();
        }

        private static void AddVideoBackground(string presentationId, string slideId)
        {
            SlidesService slidesService = GetSlideServiceClient();
            // Add a slide at index 1 using the predefined "TITLE_AND_TWO_COLUMNS" layout
            // and the ID "MyNewSlide_001".
            List<Request> requests = new List<Request>();
            Request request = new Request();

            var slideRequest = new UpdatePagePropertiesRequest();
            slideRequest.ObjectId = (slideId);
            slideRequest.Fields = "pageBackgroundFill";
            slideRequest.PageProperties = new PageProperties()
            {
                PageBackgroundFill = new PageBackgroundFill()
                {

                    StretchedPictureFill = new StretchedPictureFill()
                    {
                        ContentUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4gPTyRGSqYQoObgXWS9JRdsqBI7sQfJfgaAwQiZR43GxrR2hjWJE-Hg"
                    }
                }
            };
            request.UpdatePageProperties = slideRequest;
            //slideRequest.InsertionIndex = (1);
            //slideRequest.SlideLayoutReference = (new LayoutReference() { PredefinedLayout = "TITLE_AND_TWO_COLUMNS" });
            //request.UpdateShapeProperties = new UpdateShapePropertiesRequest()
            //    slideRequest.UpdatePageProperties = new UpdatePagePropertiesRequest() {Fields= "pageBackgroundFill" , PageProperties = new PageProperties() { PageBackgroundFill = new PageBackgroundFill() { StretchedPictureFill = new StretchedPictureFill() { ContentUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4gPTyRGSqYQoObgXWS9JRdsqBI7sQfJfgaAwQiZR43GxrR2hjWJE-Hg" } } } };
            requests.Add(request);

            // If you wish to populate the slide with elements, add create requests here,
            // using the slide ID specified above.

            // Execute the request.
            BatchUpdatePresentationRequest body =
                    new BatchUpdatePresentationRequest() { Requests = requests };
            BatchUpdatePresentationResponse response =
                    slidesService.Presentations.BatchUpdate(body, presentationId).Execute();
        }

        static string DisplayPresentation(String presentationId = "1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc")
        {
            SlidesService service = GetSlideServiceClient();

            // Define request parameters.
            PresentationsResource.GetRequest request = service.Presentations.Get(presentationId);

            // Prints the number of slides and elements in a sample presentation:
            // https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
            Presentation presentation = request.Execute();
            IList<Page> slides = presentation.Slides;
            Console.WriteLine("The presentation contains {0} slides:", slides.Count);
            for (var i = 0; i < slides.Count; i++)
            {
                var slide = slides[i];
                Console.WriteLine("- Slide #{0} contains {1} elements.", i + 1, slide.PageElements.Count);
            }
            return slides[0].ObjectId;

        }
        static string GetFirstSlideId(string presentationId )
        {
            IList<Page> slides = GetSlides(presentationId);
            return slides[0].ObjectId;

        }

        private static IList<Page> GetSlides(string presentationId)
        {
            SlidesService service = GetSlideServiceClient();

            // Define request parameters.
            PresentationsResource.GetRequest request = service.Presentations.Get(presentationId);

            // Prints the number of slides and elements in a sample presentation:
            // https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
            Presentation presentation = request.Execute();
            IList<Page> slides = presentation.Slides;
            Console.WriteLine("The presentation contains {0} slides:", slides.Count);
            for (var i = 0; i < slides.Count; i++)
            {
                var slide = slides[i];
                Console.WriteLine("- Slide #{0} contains {1} elements.", i + 1, slide.PageElements.Count);
            }

            return slides;
        }

        private static SlidesService GetSlideServiceClient()
        {
            UserCredential credential = GetCredentials();

            // Create Google Slides API service.
            var service = new SlidesService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        private static UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/slides.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        static string CreateBlankPresentation(string name)
        {
            SlidesService service = GetSlideServiceClient();
            var oobj = new Presentation() { Title = name };
           
            var presentationRequest = service.Presentations.Create(oobj);
            
            var presentation = presentationRequest.Execute();
            Console.WriteLine("Created presentation with ID:" + presentation.PresentationId);
            return presentation.PresentationId;
        }
    }
}
