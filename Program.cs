using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml.FormulaParsing;

namespace BoMBuilder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //string filePath = "D:\Users\Пользователь\Desktop\bom\test.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;            
            Console.Write("Please enter bill of materials filename:");
            string filename = Console.ReadLine();
            Console.Write("Please enter count of header's rows:");
            int headerRowsCount = Convert.ToInt32(Console.ReadLine());
            Console.Write("Please enter coun of uninformative columns:");
            int firstColumnsCount = Convert.ToInt32(Console.ReadLine());
            XLSReader reader = new XLSReader(filename);
            List<Cabinet> cabinets = reader.GenerateCabinets(headerRowsCount, firstColumnsCount);
            Console.WriteLine("Reading succesfull.");
            Console.Write("Press any key.");
            Console.ReadKey();
            Console.Write("Please enter output file template:");
            filename = Console.ReadLine();
            XLSWriter writer = new XLSWriter(filename);
            writer.GenerateXLS(cabinets);
        }

        private class XLSReader
        {
            private ExcelPackage _package;

            public XLSReader(string filename)
            {
                _package = new ExcelPackage(new FileInfo(filename));
            }

            public List<Cabinet> GenerateCabinets(int headerrowsCount, int uninformativeColumnsCount)
            {
                List<Cabinet> cabinets = new List<Cabinet>();
                List<BoMString> bomOrigin= GenerateList(headerrowsCount, uninformativeColumnsCount);
                List<string> identifiers = GetIdentifiersList(bomOrigin);
                
                foreach (string identifier in identifiers)
                {
                    Cabinet cabinet = BuildCabinet(identifier, bomOrigin);
                    cabinets.Add(cabinet); 
                }

                return cabinets;
            }

            private Cabinet BuildCabinet(string identifier, List<BoMString> bom)
            {
                Cabinet cabinet = new Cabinet(identifier);

                foreach (BoMString bomEntry in bom)
                {
                    if (bomEntry.InstallationPlace == identifier)
                    {
                        cabinet.AddString(bomEntry.Device, bomEntry.Quantity);
                    }
                }

                return cabinet;
            }

            private List<BoMString> GenerateList(int headerRowsQuantity, int uninformativeColumnsQuantity)
            {
                ExcelWorksheet worksheet = _package.Workbook.Worksheets.ElementAtOrDefault(0);
                List<BoMString> stringsList = new List<BoMString>();

                for (int row = worksheet.Dimension.Start.Row + headerRowsQuantity; row <= worksheet.Dimension.End.Row -uninformativeColumnsQuantity; row++)
                {
                    object[] fragment = new object[worksheet.Dimension.End.Column];

                    for (int column = worksheet.Dimension.Start.Column + uninformativeColumnsQuantity; column <= worksheet.Dimension.End.Column; column++)
                    {
                        fragment[column-uninformativeColumnsQuantity -1] = worksheet.Cells[row, column].GetValue<string>();
                    }

                    string descriptionOne = string.Empty;
                    string descriptionTwo = string.Empty;
                    string descriptionThree = string.Empty;
                    string numberOfType = string.Empty;
                    string orderingNumber = string.Empty;
                    string manufacturer = string.Empty;
                    string quantityUnit = string.Empty;
                    int quantity = Convert.ToInt32(fragment[7]);
                    float mass = 0;
                    string note = string.Empty;
                    string installationPlace = fragment[10].ToString();

                    if (fragment[0] != null)
                        descriptionOne = fragment[0].ToString();

                    if (fragment[1] != null)
                        descriptionTwo = fragment[1].ToString();

                    if (fragment[2] != null)
                        descriptionThree = fragment[2].ToString();

                    if (fragment[3] != null)
                        numberOfType = fragment[3].ToString();

                    if (fragment[4] != null)
                        orderingNumber = fragment[4].ToString();

                    if (fragment[5] != null)
                        manufacturer = fragment[5].ToString();

                    if (fragment[6] != null)
                        quantityUnit = fragment[6].ToString();

                    if (fragment[8] != null)
                        mass = Convert.ToSingle(fragment[8]);

                    if (fragment[9] != null)
                        note = fragment[9].ToString();

                    stringsList.Add(new BoMString(new Device(descriptionOne, descriptionTwo, descriptionThree, numberOfType,
                        orderingNumber, manufacturer, quantityUnit, mass, note), quantity, installationPlace));
                }

                return stringsList;
            }

            private List<string> GetIdentifiersList(List<BoMString> bom)
            {
                List<string> identifiers = new List<string>();

                foreach (BoMString str in bom)
                {
                    if (identifiers.Contains(str.InstallationPlace) == false)
                        identifiers.Add(str.InstallationPlace);
                }

                return identifiers;
            }
        }
        private class BoMString
        {
            private Device _device;
            private int _quantity;
            private string _installationPlace;

            public string InstallationPlace => _installationPlace;
            public int Quantity => _quantity;
            public Device Device => new Device(_device);

            public BoMString(Device device, int quantity, string installationPlace)
            {
                _device = device;
                _quantity = quantity;
                _installationPlace = installationPlace;
            }
        }

        private class XLSWriter
        {
            private ExcelPackage _package;

            public XLSWriter(string templateFilename)
            {
                _package = new ExcelPackage(new FileInfo( templateFilename));
            }

            public void GenerateXLS (List<Cabinet> cabinets, int headerRowsCount = 1, string subdirName = "output")
            {
                if (Directory.Exists(subdirName) == false)
                    Directory.CreateDirectory(subdirName);

                foreach (Cabinet cabinet in cabinets)
                {
                    List<List<string>> spredsheet = cabinet.BuildCabinetMaterialsList();
                    ExcelWorksheet worksheet = _package.Workbook.Worksheets.ElementAtOrDefault(0);
                    int rowsDelta = worksheet.Dimension.Start.Row;
                    int columnDelta = worksheet.Dimension.Start.Column;
                    
                    for (int row = 0 ; row < spredsheet.Count; row++)
                    {
                        worksheet.Cells[row + rowsDelta + headerRowsCount, columnDelta].Value = (row+rowsDelta).ToString();
                        for (int column = 0; column < spredsheet[row].Count ; column++)
                        {
                            worksheet.Cells[row + rowsDelta + headerRowsCount, column + columnDelta +1].Value = spredsheet[row][column];
                        }
                    }

                    string fileName = subdirName + '\\' + cabinet.Identifier.ToLower()+ ".xlsx";
                    _package.SaveAs(new FileInfo(fileName));
                    Console.WriteLine(fileName + " - file saved. Press any key.");
                    Console.ReadKey();

                    for (int row = worksheet.Dimension.Start.Row + headerRowsCount; row <= worksheet.Dimension.End.Row; row ++ )
                        for (int column = worksheet.Dimension.Start.Column; column <=worksheet.Dimension.End.Column; column++ )
                        worksheet.Cells[row, column].Value = null;
                }
            }
        }

        private class Cabinet
        {
            private string _identifier;
            private List<Material> _billOfMaterials;

            public string Identifier => _identifier;

            public Cabinet(string identifier)
            {
                _identifier = identifier;
                _billOfMaterials = new List<Material>();
            }

            private bool TryToAddMaterial(string orderNumber, int quantity)
            {
                bool isMaterialAlreadyExist = false;

                foreach (Material materialFromList in _billOfMaterials)
                {
                    if (materialFromList != null)
                        isMaterialAlreadyExist = materialFromList.TryToAddDevice(orderNumber, quantity);

                    if (isMaterialAlreadyExist)
                        break;
                }

                return isMaterialAlreadyExist;
            }

            public void AddString(string descriptionOne, string descriptionTwo, string descriptionThree, string numberOfType,
                string orderNumber, string manufacturer, string quantityUnit, int quantity, float mass, string note)
            {
                if (TryToAddMaterial(orderNumber, quantity) == false)
                {
                    _billOfMaterials.Add(new Material(new Device(descriptionOne, descriptionTwo, descriptionThree, numberOfType
                        , orderNumber, manufacturer, quantityUnit, mass, note), quantity));
                }
            }

            public void AddString(Device device, int quantity)
            {
                if (TryToAddMaterial(device.OrderNumber, quantity) == false)
                    _billOfMaterials.Add(new Material(device, quantity));
            }

            public List<List<string>> BuildCabinetMaterialsList()
            {
                List<List<string>> billOfMaterials = new List<List<string>>();

                foreach (Material material in _billOfMaterials)
                    billOfMaterials.Add(material.GetMaterialInfo());

                return billOfMaterials;
            }
        }

        private class Device
        {
            private string _descriptionOne;
            private string _descriptionTwo;
            private string _descriptionThree;
            private string _numberOfType;
            private string _orderNumber;
            private string _manufacturer;
            private string _quantityUnit;
            private float _mass;
            private string _additionalDescription;
            private string _description;

            public string Description => _description;
            public string NumberOfType => _numberOfType;
            public string OrderNumber => _orderNumber;
            public string Manufacturer => _manufacturer;
            public string QuantityUnit => _quantityUnit;
            public float Mass => _mass;
            public string AdditionalDescription => _additionalDescription;

            public Device(string descriptionOne, string descriptionTwo, string descriptionThree, string numberOfType, string orderNumber,
                string manufacturer, string quantityUnit, float mass, string additionalDescription, string descriptionSeparator = ". ")
            {
                _descriptionOne = descriptionOne;
                _descriptionTwo = descriptionTwo;
                _descriptionThree = descriptionThree;
                _numberOfType = numberOfType;
                _orderNumber = orderNumber;
                _manufacturer = manufacturer;
                _quantityUnit = quantityUnit;
                _mass = mass;
                _additionalDescription = additionalDescription;
                GetDescription(descriptionSeparator);
            }

            public Device(Device device, string descriptionSeparator = ". ")
            {
                _descriptionOne = device._descriptionOne;
                _descriptionTwo = device._descriptionTwo;
                _descriptionThree = device._descriptionThree;
                _numberOfType = device._numberOfType;
                _orderNumber = device._orderNumber;
                _manufacturer = device._manufacturer;
                _quantityUnit = device._quantityUnit;
                _mass = device._mass;
                _additionalDescription = device._additionalDescription;
                this.GetDescription(descriptionSeparator);
            }

            private void GetDescription(string separator)
            {
                if (_descriptionOne != string.Empty && _descriptionOne.EndsWith(separator) == false)                
                    _descriptionOne += separator;
                

                if (_descriptionTwo != string.Empty && _descriptionTwo.EndsWith(separator) == false)
                    _descriptionTwo = _descriptionTwo + separator;

                if (_descriptionThree != string.Empty && _descriptionThree.EndsWith(separator) == false)
                    _descriptionThree += separator;

                _description = _descriptionOne + _descriptionTwo + _descriptionThree;
            }
        }

        private class Material
        {
            private Device _device;
            private int _quantity;

            public int Quantity => _quantity;

            public Material(Device device, int startQuantity = 1)
            {
                _device = device;
                _quantity = startQuantity;
            }

            public Device GetDevice()
            {
                Device device = new Device(_device);
                return device;
            }

            public bool TryToAddDevice(string orderNumber, int quantity)
            {
                bool IsThisSuchDevice = _device.OrderNumber == orderNumber;

                if (IsThisSuchDevice)
                    _quantity += quantity;

                return IsThisSuchDevice;
            }

            public List<string> GetMaterialInfo()
            {
                List<string> materialInfo = new List<string>();
                materialInfo.Add(_device.Description);
                materialInfo.Add(_device.NumberOfType);
                materialInfo.Add(_device.OrderNumber);
                materialInfo.Add(_device.Manufacturer);
                materialInfo.Add(_device.QuantityUnit);
                materialInfo.Add(Convert.ToString(_quantity));
                materialInfo.Add(Convert.ToString(_device.Mass));
                materialInfo.Add(_device.AdditionalDescription);
                return materialInfo;
            }
        }
    }
}
