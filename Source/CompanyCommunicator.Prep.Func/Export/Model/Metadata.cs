// <copyright file="Metadata.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Model
{
    using System;

    /// <summary>
    /// Metadata model class.
    /// </summary>
    public class Metadata
    {
        /// <summary>
        /// Gets or sets the message title.
        /// </summary>
        public string MessageTitle { get; set; }

        /// <summary>
        /// Gets or sets the sent timestamp.
        /// </summary>
        public DateTime? SentTimeStamp { get; set; }

        /// <summary>
        /// Gets or sets the export timestamp.
        /// </summary>
        public DateTime? ExportTimeStamp { get; set; }

        /// <summary>
        /// Gets or sets the exported by user id.
        /// </summary>
        public string ExportedBy { get; set; }

        /// <summary>
        /// Gets or sets the likes count.
        /// </summary>
        public int Like { get; set; }

        /// <summary>
        /// Gets or sets the Laugh count.
        /// </summary>
        public int Laugh { get; set; }

        /// <summary>
        /// Gets or sets the Heart count.
        /// </summary>  
        public int Heart { get; set; }

        /// <summary>
        /// Gets or sets the Surprise count.
        /// </summary>  
        public int Surprise { get; set; }

        /// <summary>
        /// Gets or sets the Seen count.
        /// </summary>
        public int Seen { get; set; }

        /// <summary>
        /// Gets or sets the author.
        /// </summary> 
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Created By Name.
        /// </summary> 
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets the Succeeded count.
        /// </summary>
        public int Succeeded { get; set; }

        /// <summary>
        /// Gets or sets the Failed count.
        /// </summary>
        public int Failed { get; set; }
    }
}