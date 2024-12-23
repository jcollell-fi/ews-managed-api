﻿using global::Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.NETStandard.Core.Responses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.NETStandard.Core.Requests
{
    internal class ExportItemsRequest : MultiResponseServiceRequest<ExportItemsResponse>
    {
        private ItemIdWrapperList itemIds = new ItemIdWrapperList();

        internal ExportItemsRequest(ExchangeService exchangeService, ServiceErrorHandling serviceErrorHandling) : base(exchangeService, serviceErrorHandling)
        {

        }

        internal override ExportItemsResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new ExportItemsResponse();
        }

        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.ExportItemsResponseMessage;
        }

        internal override int GetExpectedResponseMessageCount()
        {
            return itemIds.Count;
        }

        internal override string GetXmlElementName()
        {
            return XmlElementNames.ExportItems;
        }

        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.ExportItemsResponse;
        }

        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
            itemIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.ItemIds);
        }

        public ItemIdWrapperList ItemIds
        {
            get { return this.itemIds; }
        }
    }
}