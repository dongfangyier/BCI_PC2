﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO.Ports;
using System.Windows;

namespace BCI_PC
{
    class ViewModel : INotifyPropertyChanged
    {
        private List<string> portList = new List<string>();
        public string PortName { get; set; }
        public List<string> PortList
        {
            get => portList;
            set
            {
                portList = value;
                NotifyPropertyChanged("PortList");
            }
        }
        public BCIEntity bCIEntity;

        public ViewModel()
        {
            BCIEntity.savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            BCIEntity.initSave();
            bCIEntity = new BCIEntity(PortName);
            bCIEntity.StartBCI();
            bCIEntity.OpenChannel();
        }

       


        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}
