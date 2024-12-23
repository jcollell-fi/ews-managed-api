﻿using global::Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.NETStandard.Core.ServiceObjects.Schemas
{
    /// <summary>
    /// Upload Schema
    /// </summary>
    [Schema]
    public class UploadSchema : ServiceObjectSchema
    {
        private static class FieldUris
        {
            public const string ItemId = "upload:ItemId";
            public const string ParentFolderId = "upload:ParentFolderId";
            public const string Data = "upload:Data";
        }


        /// <summary>
        /// Defines the uploads id property
        /// </summary>
        public static readonly PropertyDefinition Id = new ComplexPropertyDefinition<ItemId>(
            XmlElementNames.ItemId,
            FieldUris.ItemId,
            PropertyDefinitionFlags.CanSet,
            ExchangeVersion.Exchange2010,
            delegate () { return new ItemId(); }
            );
        /// <summary>
        /// Defines the uploads ParentFolderId property.
        /// </summary>
        public static readonly PropertyDefinition ParentFolderId = new ComplexPropertyDefinition<FolderId>(
            XmlElementNames.ParentFolderId,
            FieldUris.ParentFolderId,
            PropertyDefinitionFlags.CanSet,
            ExchangeVersion.Exchange2010,
            delegate () { return new FolderId(); }
            );

        /// <summary>
        /// Defines the data property.
        /// </summary>
        public static readonly PropertyDefinition Data = new ByteArrayPropertyDefinition(
            XmlElementNames.Data,
            FieldUris.Data,
            PropertyDefinitionFlags.CanSet,
            ExchangeVersion.Exchange2010
            );

        internal static readonly UploadSchema Instance = new UploadSchema();

        internal override void RegisterProperties()
        {
            base.RegisterProperties();
            this.RegisterProperty(ParentFolderId);
            this.RegisterProperty(Id);
            this.RegisterProperty(Data);
        }

        internal UploadSchema() : base() { }
    }
}
