using System;

namespace SharePoint_RER_AzureFunction
{
    public class SharePointEventResponse
    {
        public string Status { get; set; }
        public string ErrorMessage { get; set; }
        public bool CancelEvent { get; set; }
    }
}