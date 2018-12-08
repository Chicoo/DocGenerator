using System;
using System.Runtime.Serialization;
namespace DocumentGenerator.Common
{
    /// <summary>
    /// Represents an exception that is thrown when no file name is provided for the document.
    /// </summary>
    [Serializable()]
    public class FilenameEmptyException : ApplicationException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.FilenameEmptyException"/> class.
        /// </summary>
        public FilenameEmptyException() : base() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.FilenameEmptyException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The error message that explains the error.</param>
        public FilenameEmptyException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.FilenameEmptyException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The error message that explains the error.</param>
        /// <param name="inner">The exception that is the cause of the current exceprion. If the <b>inner</b> parameter is not null, the current exception is raised in a catch block that handles the inner exception.</param>
        public FilenameEmptyException(string message, System.Exception inner) : base(message, inner) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.FilenameEmptyException"/> class.
        /// </summary>
        private FilenameEmptyException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Represents an exception that is thrown when a duplicate custom tag is added.
    /// </summary>
    [Serializable()]
    public class DuplicateCustomTagException : ApplicationException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.DuplicateCustomTagException"/> class.
        /// </summary>
        public DuplicateCustomTagException() : base() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.DuplicateCustomTagException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The error message that explains the error.</param>
        public DuplicateCustomTagException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.DuplicateCustomTagException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The error message that explains the error.</param>
        /// <param name="inner">The exception that is the cause of the current exceprion. If the <b>inner</b> parameter is not null, the current exception is raised in a catch block that handles the inner exception.</param>
        public DuplicateCustomTagException(string message, System.Exception inner) : base(message, inner) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator.Common.DuplicateCustomTagException"/> class.
        /// </summary>
        private DuplicateCustomTagException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

}
