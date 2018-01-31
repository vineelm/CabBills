namespace CabBills
{
    class BillData
    {
        public string Date { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string BaseFare { get; set; }
        public string DistanceKms { get; set; }
        public string DistanceFare { get; set; }
        public string RideTime { get; set; }
        public string RideFare { get; set; }

        public string AdvanceBooking { get; set; }
        public string Taxes { get; set; }
        public string TotalBill { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string SourceAddress { get; set; }
        public string DestinationAddress { get; set; }
    }
}
