using System;
using InTheHand.Net;



namespace cAtod_2
{
    internal class Device
    {
        public String blName;
        public BluetoothAddress blAddress;
        public BluetoothEndPoint blEndpoint;

        public Device(string blName, BluetoothAddress blthAddress, BluetoothEndPoint blEndPoint) 
        { 
            this.blName = blName;
            this.blAddress = blthAddress;   
            this.blEndpoint = blEndPoint;
        }
    }

    internal class Parsel
    {
        public string mode;
        public double borehole;
        public double? jobTime = null;
        public double? pauseTime = null;
    }
}
