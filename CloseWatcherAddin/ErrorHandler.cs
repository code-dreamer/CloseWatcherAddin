using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace CloseWatcherAddin
{
	/// <summary>
	/// Some util methods to process errors.
	/// </summary>
	static class ErrorHandler
	{
		/// <summary>
		///	Show exception error to user.
		/// </summary>
		public static void ReportError(Exception exception)
		{
			string errorMsg = exception == null ? "Unknown error" : exception.Message;
			ReportError(errorMsg);
		}

		/// <summary>
		///	Show string error to user.
		/// </summary>
		public static void ReportError(string message)
		{
			Debug.Assert(!String.IsNullOrEmpty(message));

			MessageBox.Show(message, FormatCaption(), MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		/// <summary>
		///	Format msg box error caption from product name 
		/// from assembly.
		/// </summary>
		private static string FormatCaption()
		{
			var attribute = Assembly.GetExecutingAssembly().GetAssemblyAttribute<AssemblyProductAttribute>();
			if (attribute != null) 
				return attribute.Product;
			
			Debug.Fail("AssemblyProductAttribute is missing");
			return "Error";
		}

		/// <summary>
		///	Extension function for retrieving special attribute 
		/// from assembly.
		/// </summary>
		/// <param name='assembly'>
		/// Required assembly.
		/// </param>
		public static T GetAssemblyAttribute<T>(this Assembly assembly) 
			where T :  Attribute
		{
			object[] attributes = assembly.GetCustomAttributes(typeof(T), false);
			return attributes.Length == 0 ? null : attributes.OfType<T>().SingleOrDefault();
		}

	}
}
