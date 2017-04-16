using System;
using System.Collections.Generic;
using Windows.ApplicationModel.Appointments;
using Windows.Foundation;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace AddToCalendar
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public static MainPage Current;
        private MainPage rootPage = MainPage.Current;
        StorageFile file;

        public MainPage()
        {
            this.InitializeComponent();
        }

        private async void BtnFile_ClickAsync(object sender, RoutedEventArgs e)
        {
            FileOpenPicker openPicker = new FileOpenPicker();
            openPicker.ViewMode = PickerViewMode.Thumbnail;
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            openPicker.FileTypeFilter.Add(".txt");
            openPicker.FileTypeFilter.Add(".csv");

            file = await openPicker.PickSingleFileAsync();
            if (file != null)
            {
                txtStatus.Text = "Picked file: " + file.Name +"\n";
            }
            else
            {
                txtStatus.Text = "Operation cancelled.\n";
            }
        }

        private async void BtnAppnt_ClickAsync(object sender, RoutedEventArgs e)
        {
            var rect = GetElementRect(sender as FrameworkElement);

            // Store contents of the file into an enumerable List.
            IList<string> lines = await FileIO.ReadLinesAsync(file);
            char chDelimiter;

            // Choose delimited
            if (file.Name.Contains(".csv"))
                chDelimiter = ',';
            else
                chDelimiter = ';';

            foreach (var line in lines)
            {
                // Skip the header or Empty Line
                if (line.Contains("Subject,Date,StartTime,Duration") || line.Equals(String.Empty))
                    continue;

                // Split the contents of the line based on the delimiter
                string[] appointmentDetails = line.Split(chDelimiter);

                // Create a new Appointment object
                Appointment appointment = new Appointment();

                // Expected format of string: Subject,Date,StartTime,Duration
                appointment.Subject = appointmentDetails[0];
                DateTime dateTime = DateTime.ParseExact(String.Format("{0} {1}", appointmentDetails[1], appointmentDetails[2]), 
                                                        "dd-MM-yyyy HH:mm", 
                                                        System.Globalization.CultureInfo.InvariantCulture);
                DateTimeOffset startTime = new DateTimeOffset(dateTime);
                appointment.StartTime = startTime;
                appointment.Duration = TimeSpan.FromHours(float.Parse(appointmentDetails[3]));
                appointment.Reminder = TimeSpan.FromHours(1);

                String appointmentId = await AppointmentManager.ShowAddAppointmentAsync(appointment,rect);
                
                // Check if appointment was added successfully
                if (appointmentId != String.Empty)
                {
                    if (appointmentId.Length > 10)
                        appointmentId = appointmentId.Substring(0,10);
                    txtStatus.Text = "Appointment Id: " + appointmentId;
                }
                else
                {
                    txtStatus.Text = "Appointment Not added";
                }
            }
        }

        // Helper function to calculate an element's rectangle in root-relative coordinates.
        public static Rect GetElementRect(Windows.UI.Xaml.FrameworkElement element)
        {
            Windows.UI.Xaml.Media.GeneralTransform transform = element.TransformToVisual(null);
            Point point = transform.TransformPoint(new Point());
            return new Rect(point, new Size(element.ActualWidth, element.ActualHeight));
        }
    }
}
