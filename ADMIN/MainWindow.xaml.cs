using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using Microsoft.Win32;
using System.Data;
using LiveChartsCore.SkiaSharpView.Extensions;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.VisualElements;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;
using System.Collections.ObjectModel;
using System.Windows.Media.Media3D;
using LiveChartsCore.Measure;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using System.Runtime.Remoting.Contexts;
using LiveChartsCore;
using System.ComponentModel;

namespace ADMIN
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private Role _currentRole = new Role();
        private Country _currentCountry = new Country();
        //private Ticket _currentTicket = new Ticket();
        private Class_ticket _currentClass_ticket = new Class_ticket();
        private Age _currentAge = new Age();
        private Regular_plane _currentRegular_plane = new Regular_plane();
        private Airplane _currentAirplane = new Airplane();
        private Flight _currentFlight = new Flight();
        private Airline _currentAirline = new Airline();
        private Airport _currentAirport = new Airport();
        private History_tickets _currentHistory_tickets = new History_tickets();

        private DataGrid dataGrid;

        public MainWindow()
        {
            InitializeComponent();

        }



        ///<summary>
        /// Переменная отвечающая за наполнение графика
        /// </summary>
        public List<ISeries<int>> Series { get; set; } = new List<ISeries<int>>();

        //Datagrid

        ///<summary>
        /// Метод загрузки датагрида при видимости
        /// </summary>
        private void dtgRole_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentRole;
                dtgRole.ItemsSource = AirplaneEntities.GetContext().Roles.ToList();
                dataGrid = dtgRole;
            }
        }
        private void dtgClassTicket_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentClass_ticket;
                dtgClassTicket.ItemsSource = AirplaneEntities.GetContext().Class_tickets.ToList();
                dataGrid = dtgClassTicket;
            }
        }
        private void dtgAge_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentAge;
                dtgAge.ItemsSource = AirplaneEntities.GetContext().Ages.ToList();
                dataGrid = dtgAge;
            }
        }
        private void dtgAirlines_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentAirline;
                dtgAirlines.ItemsSource = AirplaneEntities.GetContext().Airlines.ToList();
                dataGrid = dtgAirlines;
            }
        }
        private void dtgAirport_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentAirport;
                dtgAirport.ItemsSource = AirplaneEntities.GetContext().Airports.ToList();
                dataGrid = dtgAirport;
            }
        }
        private void dtgFlight_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentFlight;
                dtgFlight.ItemsSource = AirplaneEntities.GetContext().Flights.ToList();
                dataGrid = dtgFlight;
            }
        }
        private void dtgRegularPlane_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentRegular_plane;
                dtgRegularPlane.ItemsSource = AirplaneEntities.GetContext().Regular_planes.ToList();
                dataGrid = dtgRegularPlane;
            }
        }
        private void dtgAirplane_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentAirplane;
                dtgAirplane.ItemsSource = AirplaneEntities.GetContext().Airplanes.ToList();
                dataGrid = dtgAirplane;
            }
        }
        private void dtgHistory_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentHistory_tickets;
                dtgHistory.ItemsSource = AirplaneEntities.GetContext().History_tickets.ToList();
                dataGrid = dtgHistory;
            }
        }
        private void dtgView_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                dtgView.ItemsSource = AirplaneEntities.GetContext().SCHEDULEs.ToList();
                dataGrid = dtgView;
            }
        }
        private void dtgCountry_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                DataContext = _currentCountry;
                dtgCountry.ItemsSource = AirplaneEntities.GetContext().Countries.ToList();
                dataGrid = dtgCountry;
            }
        }

        //ComboBox

        ///<summary>
        /// Метод загрузки выпадающего списка
        /// </summary>
        private void cmbNameAirlines_Loaded(object sender, RoutedEventArgs e)
        {
            cmbNameAirlines.ItemsSource = AirplaneEntities.GetContext().Airlines.ToList();
        }
        private void cmbRegularPlane_Loaded(object sender, RoutedEventArgs e)
        {
            cmbRegularPlane.ItemsSource = AirplaneEntities.GetContext().Regular_planes.ToList();
        }
        private void cmbAirport_Loaded(object sender, RoutedEventArgs e)
        {
            cmbAirport.ItemsSource = AirplaneEntities.GetContext().Airports.ToList();
        }
        private void cmbAirport1_Loaded(object sender, RoutedEventArgs e)
        {
            cmbAirport1.ItemsSource = AirplaneEntities.GetContext().Airports.ToList();
        }
        private void cmbAirplane_Loaded(object sender, RoutedEventArgs e)
        {
            cmbAirplane.ItemsSource = AirplaneEntities.GetContext().Airplanes.ToList();
        }
        private void cmbCountry_Loaded(object sender, RoutedEventArgs e)
        {
            cmbCountry.ItemsSource = AirplaneEntities.GetContext().Countries.ToList();
        }
        private void cmbFilterAirlines_Loaded(object sender, RoutedEventArgs e)
        {
            cmbFilterAirlines.ItemsSource = AirplaneEntities.GetContext().Airlines.ToList();
        }

        //Country

        ///<summary>
        /// Метод добавления и редактирования записи
        /// </summary>
        private void btnCountrySave_Click(object sender, RoutedEventArgs e)
        {
            var countryIsInsertOrUpdate = dtgCountry.SelectedItems.Cast<Country>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtCountry.Text))
                errors.AppendLine("Укажите название страны!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (countryIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentCountry.Country1 = txtCountry.Text;
                AirplaneEntities.GetContext().Countries.Add(_currentCountry);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgCountry.ItemsSource = AirplaneEntities.GetContext().Countries.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        ///<summary>
        /// Метод удаления записи
        /// </summary>
        private void btnCountryDelete_Click(object sender, RoutedEventArgs e)
        {
            var countryIsRemoving = dtgCountry.SelectedItems.Cast<Country>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {countryIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Countries.RemoveRange(countryIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgCountry.ItemsSource = AirplaneEntities.GetContext().Countries.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }


        //Role
        private void btnRoleSave_Click(object sender, RoutedEventArgs e)
        {
            var roleIsInsertOrUpdate = dtgRole.SelectedItems.Cast<Role>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtRole.Text))
                errors.AppendLine("Укажите название роли!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (roleIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentRole.Name_role = txtRole.Text;
                AirplaneEntities.GetContext().Roles.Add(_currentRole);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgRole.ItemsSource = AirplaneEntities.GetContext().Roles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }
        private void btnRoleDelete_Click(object sender, RoutedEventArgs e)
        {
            var roleIsRemoving = dtgRole.SelectedItems.Cast<Role>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {roleIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Roles.RemoveRange(roleIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgRole.ItemsSource = AirplaneEntities.GetContext().Roles.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Class_ticket
        private void btnClassSave_Click(object sender, RoutedEventArgs e)
        {
            var classIsInsertOrUpdate = dtgClassTicket.SelectedItems.Cast<Class_ticket>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtClass.Text))
                errors.AppendLine("Укажите название класс!");
            if (string.IsNullOrWhiteSpace(txtFactorClass.Text))
                errors.AppendLine("Укажите коэфицент класса!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (classIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentClass_ticket.Name_class = txtClass.Text;
                _currentClass_ticket.Factor_class = Convert.ToDecimal(txtFactorClass.Text);
                AirplaneEntities.GetContext().Class_tickets.Add(_currentClass_ticket);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgRole.ItemsSource = AirplaneEntities.GetContext().Roles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnClassDelete_Click(object sender, RoutedEventArgs e)
        {
            var classIsRemoving = dtgClassTicket.SelectedItems.Cast<Class_ticket>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {classIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Class_tickets.RemoveRange(classIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgClassTicket.ItemsSource = AirplaneEntities.GetContext().Class_tickets.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }


        ///<summary>
        /// Использование хранимых процедур
        /// </summary>
        //Age
        private void btnAgeSave_Click(object sender, RoutedEventArgs e)
        {
            var ageInsertOrUpdate = dtgAge.SelectedItems.Cast<Age>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtRange.Text))
                errors.AppendLine("Укажите название класс!");
            if (string.IsNullOrWhiteSpace(txtFactorAge.Text))
                errors.AppendLine("Укажите коэфицент класса!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (ageInsertOrUpdate.FirstOrDefault() == null)
            {
                AirplaneEntities.GetContext().INSERT_AGE(Convert.ToDecimal(txtFactorAge.Text), txtRange.Text);
            }

            if (ageInsertOrUpdate.FirstOrDefault() != null)
            {
                AirplaneEntities.GetContext().UPDATE_AGE(ageInsertOrUpdate.FirstOrDefault().ID_Age, Convert.ToDecimal(txtFactorAge.Text), txtRange.Text);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgRole.ItemsSource = AirplaneEntities.GetContext().Roles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnAgeDelete_Click(object sender, RoutedEventArgs e)
        {
            var ageIsRemoving = dtgAge.SelectedItems.Cast<Age>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {ageIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().DELETE_AGE(ageIsRemoving.FirstOrDefault().Range);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgAge.ItemsSource = AirplaneEntities.GetContext().Ages.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Airlines
        private void btnAirlinesSave_Click(object sender, RoutedEventArgs e)
        {
            var airlinesIsInsertOrUpdate = dtgAirlines.SelectedItems.Cast<Airline>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtNameAirlines.Text))
                errors.AppendLine("Укажите название авиакомпании!");
            if (string.IsNullOrWhiteSpace(txtLWPP.Text))
                errors.AppendLine("Укажите багаж/человек!");
            if (string.IsNullOrWhiteSpace(txtLuggage_price.Text))
                errors.AppendLine("Укажите цену!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (airlinesIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentAirline.Name_airlines = txtNameAirlines.Text;
                _currentAirline.LWPP = Convert.ToInt32(txtLWPP.Text);
                _currentAirline.Luggage_price = Convert.ToDecimal(txtLuggage_price.Text);
                AirplaneEntities.GetContext().Airlines.Add(_currentAirline);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgAirlines.ItemsSource = AirplaneEntities.GetContext().Airlines.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnAirlinesDelete_Click(object sender, RoutedEventArgs e)
        {
            var airlinesIsRemoving = dtgAirlines.SelectedItems.Cast<Airline>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {airlinesIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Airlines.RemoveRange(airlinesIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgAirlines.ItemsSource = AirplaneEntities.GetContext().Airlines.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Airport
        private void btnAirportSave_Click(object sender, RoutedEventArgs e)
        {
            var airportIsInsertOrUpdate = dtgAirport.SelectedItems.Cast<Airport>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtNameAirport.Text))
                errors.AppendLine("Укажите название аэропорта!");
            if (string.IsNullOrWhiteSpace(cmbCountry.SelectedValue.ToString()))
                errors.AppendLine("Укажите страну!");
            if (string.IsNullOrWhiteSpace(txtLocation.Text))
                errors.AppendLine("Укажите адрес!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (airportIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentAirport.Airport_name = txtNameAirport.Text;
                _currentAirport.Country_ID = Convert.ToInt32(cmbCountry.SelectedValue);
                _currentAirport.Location = txtLocation.Text;
                AirplaneEntities.GetContext().Airports.Add(_currentAirport);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgAirport.ItemsSource = AirplaneEntities.GetContext().Airports.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnAirportDelete_Click(object sender, RoutedEventArgs e)
        {
            var airportIsRemoving = dtgAirport.SelectedItems.Cast<Airport>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {airportIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Airports.RemoveRange(airportIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgAirport.ItemsSource = AirplaneEntities.GetContext().Airports.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Flight
        private void btnFlightSave_Click(object sender, RoutedEventArgs e)
        {
            var airportIsInsertOrUpdate = dtgFlight.SelectedItems.Cast<Flight>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(dtpDepatureDateandTime.Text))
                errors.AppendLine("Укажите дату и время отправления!");
            if (string.IsNullOrWhiteSpace(dtpArrivalDateandTime.Text))
                errors.AppendLine("Укажите дату и время прилёта!");
            if (string.IsNullOrWhiteSpace(txtFlightPrice.Text))
                errors.AppendLine("Укажите адрес!");
            if (string.IsNullOrWhiteSpace(cmbNameAirlines.SelectedValue.ToString()))
                errors.AppendLine("Укажите название аэропорта!");
            if (string.IsNullOrWhiteSpace(cmbRegularPlane.SelectedValue.ToString()))
                errors.AppendLine("Укажите вместимость!");
            if (string.IsNullOrWhiteSpace(cmbRegularPlane.SelectedValue.ToString()))
                errors.AppendLine("Укажите  аэропорт отбытия!!");
            if (string.IsNullOrWhiteSpace(cmbAirport.SelectedValue.ToString()))
                errors.AppendLine("Укажите аэропорт прибытия!");


            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (airportIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentFlight.Depature_date_and_time = Convert.ToDateTime(dtpDepatureDateandTime.Value);
                _currentFlight.Arrival__date_and_time = Convert.ToDateTime(dtpArrivalDateandTime.Value);
                _currentFlight.Flight_price = Convert.ToDecimal(txtFlightPrice.Text);
                _currentFlight.Airlines_ID = Convert.ToInt32(cmbNameAirlines.SelectedValue);
                _currentFlight.Regular_plane_ID = Convert.ToInt32(cmbRegularPlane.SelectedValue);
                _currentFlight.Depature_location_ID = Convert.ToInt32(cmbAirport.SelectedValue);
                _currentFlight.Arrival_location_ID = Convert.ToInt32(cmbAirport1.SelectedValue);
                AirplaneEntities.GetContext().Flights.Add(_currentFlight);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgFlight.ItemsSource = AirplaneEntities.GetContext().Flights.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnFlightDelete_Click(object sender, RoutedEventArgs e)
        {
            var flightIsRemoving = dtgFlight.SelectedItems.Cast<Flight>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {flightIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Flights.RemoveRange(flightIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgFlight.ItemsSource = AirplaneEntities.GetContext().Flights.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //RegularPlane
        private void btnRegularPlaneSave_Click(object sender, RoutedEventArgs e)
        {
            var regularPlaneInsertOrUpdate = dtgRegularPlane.SelectedItems.Cast<Regular_plane>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtCapacity.Text))
                errors.AppendLine("Укажите название аэропорта!");
            if (string.IsNullOrWhiteSpace(cmbAirplane.Text.ToString()))
                errors.AppendLine("Укажите страну!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (regularPlaneInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentRegular_plane.Capacity = Convert.ToInt32(txtCapacity.Text);
                _currentRegular_plane.Airplane_ID = Convert.ToInt32(cmbAirplane.SelectedValue);
                _currentAirport.Location = txtLocation.Text;
                AirplaneEntities.GetContext().Regular_planes.Add(_currentRegular_plane);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgRegularPlane.ItemsSource = AirplaneEntities.GetContext().Regular_planes.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnRegularDelete_Click(object sender, RoutedEventArgs e)
        {
            var regularPlaneIsRemoving = dtgRegularPlane.SelectedItems.Cast<Regular_plane>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {regularPlaneIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Regular_planes.RemoveRange(regularPlaneIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgRegularPlane.ItemsSource = AirplaneEntities.GetContext().Regular_planes.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Airplane
        private void btnAirplaneSave_Click(object sender, RoutedEventArgs e)
        {
            var airplaneIsInsertOrUpdate = dtgAirplane.SelectedItems.Cast<Airplane>();

            StringBuilder errors = new StringBuilder();
            if (string.IsNullOrWhiteSpace(txtBrandandNumber.Text))
                errors.AppendLine("Укажите дату и время отправления!");



            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }


            if (airplaneIsInsertOrUpdate.FirstOrDefault() == null)
            {
                _currentAirplane.Brand_and_number = txtBrandandNumber.Text;
                AirplaneEntities.GetContext().Airplanes.Add(_currentAirplane);
            }

            try
            {
                AirplaneEntities.GetContext().SaveChanges();
                MessageBox.Show("Успешно!");
                dtgAirplane.ItemsSource = AirplaneEntities.GetContext().Airplanes.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnAirplaneDelete_Click(object sender, RoutedEventArgs e)
        {
            var airplaneIsRemoving = dtgAirplane.SelectedItems.Cast<Airplane>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие {airplaneIsRemoving.Count} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    AirplaneEntities.GetContext().Airplanes.RemoveRange(airplaneIsRemoving);
                    AirplaneEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!!!");

                    dtgRegularPlane.ItemsSource = AirplaneEntities.GetContext().Regular_planes.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        //Search
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                if (cmbFilterAirlines.SelectedValue != null)
                {
                    dtgView.ItemsSource = AirplaneEntities.GetContext().SCHEDULEs.Where(x => x.Место_прилёта.Contains(txtSearch.Text)).Where(x => x.Авиакомпания == cmbFilterAirlines.SelectedValue.ToString()).ToList();
                }
                else
                {
                    dtgView.ItemsSource = AirplaneEntities.GetContext().SCHEDULEs.Where(x => x.Место_прилёта.Contains(txtSearch.Text)).ToList();
                }
            }
            else if (cmbFilterAirlines.SelectedValue != null)
            {
                dtgView.ItemsSource = AirplaneEntities.GetContext().SCHEDULEs.Where(x => x.Авиакомпания == cmbFilterAirlines.SelectedValue.ToString()).ToList();
            }
            else
            {
                dtgView.ItemsSource = AirplaneEntities.GetContext().SCHEDULEs.ToList();
            }
        }

        //BackUp
        ///<summary>
        /// Метод выгрузка базы данных в sql
        /// </summary>
        private void btnBackUp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AirplaneEntities.GetContext().Database.ExecuteSqlCommand(TransactionalBehavior.DoNotEnsureTransaction, $"BACKUP DATABASE Airplane TO DISK = 'C:\\Program Files\\Microsoft SQL Server\\MSSQL16.MSSQLSERVER\\MSSQL\\Backup\\Airplane.bak'  WITH FORMAT, MEDIANAME = 'SQLServerBackups', NAME = '{"Backup_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}'");
                MessageBox.Show("Успешно!!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        //Export
        ///<summary>
        /// Метод выгрузки записей в csv
        /// </summary>
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < dataGrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = dataGrid.Columns[j].Header;
            }
            for (int i = 0; i < dataGrid.Columns.Count; i++)
            {
                for (int j = 0; j < dataGrid.Items.Count; j++)
                {
                    TextBlock b = dataGrid.Columns[i].GetCellContent(dataGrid.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
                }
            }
        }

        ///<summary>
        /// Наполнение графика данными
        /// </summary>

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            var flight = AirplaneEntities.GetContext().Flights.OrderBy(p => p.Airline.Name_airlines).GroupBy(p => p.Airline.Name_airlines).ToList();

            for (int i = 0; i < flight.Count; i++)
            {
                var ser = new PieSeries<int> { Values = new int[] { flight[i].Count() }, Name = flight[i].Key };
                Series.Add(ser);
            }

            Pie.Series = Series;

        }
    }
}
