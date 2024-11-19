using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.NETStandard.Core.ServiceObjects.Schemas;
using Microsoft.Exchange.WebServices.NETStandard.Enumerations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.NETStandard.Core.ServiceObjects.Items
{
    

    namespace Microsoft.Exchange.WebServices.Data
    {
        /// <summary>
        /// Represents and item to be uploaded to exchange.
        /// </summary>
        public class UploadItem : ServiceObject
        {
            /// <summary>
            /// The default contructor for UploadItem
            /// </summary>
            /// <param name="service">The exchange service this item is related to.</param>
            public UploadItem(ExchangeService service) : base(service) { }

            internal override ServiceObjectSchema GetSchema()
            {
                return UploadSchema.Instance;
            }

            internal override ExchangeVersion GetMinimumRequiredServerVersion()
            {
                return ExchangeVersion.Exchange2010;
            }

            internal override PropertyDefinition GetIdPropertyDefinition()
            {
                return UploadSchema.Id;
            }

            internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(PropertySet propertySet, CancellationToken token)
            {
                throw new NotImplementedException();
            }

            internal override Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(DeleteMode deleteMode, SendCancellationsMode? sendCancellationsMode, AffectedTaskOccurrence? affectedTaskOccurrences, CancellationToken token)
            {
                throw new NotImplementedException();
            }

            /// <summary>
            /// The Id of the item to upload
            /// </summary>
            public ItemId Id
            {
                get { return (ItemId)this.PropertyBag[GetIdPropertyDefinition()]; }
                set { this.PropertyBag[GetIdPropertyDefinition()] = value; }
            }

            /// <summary>
            /// The id of the folder to upload the item to.
            /// </summary>
            public FolderId ParentFolderId
            {
                get { return (FolderId)this.PropertyBag[UploadSchema.ParentFolderId]; }
                set { this.PropertyBag[UploadSchema.ParentFolderId] = value; }
            }

            /// <summary>
            /// The data of the item.
            /// </summary>
            public byte[] Data
            {
                get { return (byte[])this.PropertyBag[UploadSchema.Data]; }
                set { this.PropertyBag[UploadSchema.Data] = value; }
            }

            /// <summary>
            /// The action to take for the uploaded item.
            /// </summary>
            public CreateAction CreateAction { get; set; }
        }
    }
}
