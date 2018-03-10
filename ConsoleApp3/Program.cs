using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ConsoleApp3
{
	static class Program
	{
		static void Main(string[] args)
		{
			Outlook.Application OutlookApp = null;
			if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Length > 0)
			{

				// If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
				OutlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
			}
			else
			{

				// If not, create a new instance of Outlook and log on to the default profile.
				OutlookApp = new Outlook.Application();
			}

			string defaultProfile = OutlookApp.DefaultProfileName;
			Outlook.NameSpace nameSpace = OutlookApp.GetNamespace("MAPI");
			nameSpace.Logon(Profile: "", Password: "", ShowDialog: Missing.Value, NewSession: Missing.Value);
			var OutlookSession = OutlookApp.Session;
			Stores stores = nameSpace.Stores;

			//Task.Factory.StartNew(() => EnumerateFolderItems(stores, "2007-00-00", "2007-00-00"));
			EnumerateFolderItems(stores, "2007-00-00", "2007-00-00");
			//EnumerateStores(stores);
			//DeleteEmptyFolders(stores);
			//for (int i = 0; i < 10; i++)
			//{
			//	EnumerateItems(stores);
			//}
			DeleteEmptyFolders(stores);

			//CreateSortStores(OutlookApp);
			//NameSortStores(OutlookApp);
		}

		private static void EnumerateFolderItems(Stores s, string v1, string v2)
		{
			Store S = s[v1];
			Folder F = GetFolder(S, v2);
			EnumerateItems(F, 0);
		}

		private static void EnumerateItems(Stores stores)
		{

			foreach (Store s in stores)
			{
				System.Diagnostics.Debug.WriteLine(String.Format("{0} {1}",
					s.DisplayName,
					s.FilePath
					));

				Folder f = s.GetRootFolder() as Outlook.Folder;
				if (f.DefaultItemType != OlItemType.olMailItem)
				{
					continue;
				}
				EnumerateItems(f.Folders, 0);
			}
		}

		private static void EnumerateItems(Folders F, int Depth)
		{
			foreach (Folder f in F)
			{
				if (f.Name.EndsWith("Deleted Items"))
				{
					continue;
				}
				if (f.DefaultItemType != OlItemType.olMailItem)
				{
					continue;
				}
				if (f.Name.EndsWith("-00"))
				{
					if (f.Name.EndsWith("-00-00"))
					{ }
					else
						continue;
				}

				if (f.Folders.Count > 0)
				{
					EnumerateItems(f.Folders, Depth + 1);
				}

				if (f.Items.Count > 0)
				{
					EnumerateItems(f, Depth);
				}
			}

		}

		private static void EnumerateItems(Folder f, int Depth)
		{
			Items itms = f.Items;
			var n = f.Items.Count;

			//System.Diagnostics.Debug.WriteLine(String.Format("EnumerateItems() {0} {1} {2}",
			//	n,
			//	Depth,
			//	f.FullFolderPath));

			int i = 0;
			foreach (var itm in itms)
			{
				Itemiz(itm, f, Depth);

				System.Diagnostics.Debug.Write(String.Format("Itm {0} of {1}",
				i,
				n
				));
				if (i > 1000)
				{
					return;
				}

				i++;
			}
		}

		private static void Itemiz(object itm, Folder f, object depth)
		{
			Outlook.MailItem m = itm as Outlook.MailItem;
			if (m == null)
			{
				return;
			}

			DateTime rcvd;
			try
			{
				rcvd = new DateTime(m.ReceivedTime.Ticks);
			}
			catch (System.Exception e)
			{
				System.Diagnostics.Debug.WriteLine(String.Format("Date Conversion exception {0}",
					e));
				return;
			}

			String Year = rcvd.Year.ToString("0000");
			String Month = rcvd.Month.ToString("00");

			Store StoresFrom = f.Store;
			Store StoresTo = GetOrMakeYearStore(f.Application, rcvd.Year);

			Folder inbox = GetOrMakeFolder(StoresTo, "Inbox");

			String YearString = String.Format(@"{0}-00-00", Year);
			Folder YearFolder = GetOrMakeFolder(inbox, YearString);

			Folder MonthFolder = GetOrMakeFolder(YearFolder, String.Format(@"{0}-{1}-00", Year, Month));
			MonthFolder.ShowItemCount = OlShowItemCount.olShowTotalItemCount;

			System.Diagnostics.Debug.WriteLine(String.Format("<{0}><{4}> {1} {2} ==> <{3}>",
				StoresFrom.DisplayName,
				Year,
				Month,
				StoresTo.DisplayName, f.Name
				));

			m.UnRead = true;
			MailItem mm = m.Move(MonthFolder) as MailItem;
			//Explorer p = MonthFolder.Application.ActiveExplorer();
			//p.CurrentFolder = f;
			//p.ShowPane(OlPane.olFolderList, true);
			//p.ShowPane(OlPane.olPreview, false);
			//p.ShowPane(OlPane.olOutlookBar, false);
		}

		private static Store GetOrMakeYearStore(Application A, int year)
		{
			Outlook.NameSpace n = A.GetNamespace("MAPI");
			Stores Stores = n.Stores;

			String yearStore = String.Format("{0}-00-00", year);
			Store StoresTo = null;
			try
			{
				StoresTo = Stores[yearStore];
			}
			catch (System.Exception e)
			{

				System.Diagnostics.Debug.WriteLine(String.Format("Liability Blabber here {0} ",
				e
				));
				goto jumper;
			}
			return StoresTo;
			jumper:;
			return CreateYearStore(n, year);
		}
		public static Outlook.Store CreateYearStore(NameSpace Na, int year)
		{

			String PrePath = @"L:\\Outlook on L\";
			String PST_displayName = String.Format("{0}-00-00", year.ToString("0000"));
			String PST_name = String.Format(@"{0}\{1}.pst", PrePath, PST_displayName);
			Na.AddStoreEx(PST_name, OlStoreType.olStoreUnicode);
			Store St = GetStore(Na, PST_displayName);
			Folder rF = St.GetRootFolder() as Folder;
			rF.Name = PST_displayName;
			return St;

		}
		public static Outlook.Folder GetOrMakeFolder(Store S, String Fn)
		{
			return GetOrMakeFolder(S.GetRootFolder() as Outlook.Folder, Fn);
		}

		public static Outlook.Folder GetOrMakeFolder(Folder F, String Fn)
		{
			if (F == null)
			{
				return null as Folder;
			}

			Folder ff = GetFolder(F, Fn);
			if (ff != null)
			{
				return ff;
			}

			Folder fff = F.Folders.Add(Fn, OlDefaultFolders.olFolderInbox) as Outlook.Folder;
			return fff;
		}
		private static void DeleteEmptyFolders(Stores stores)
		{
			foreach (Store s in stores)
			{
				if (s.FilePath.EndsWith("-00-00.pst"))
				{
					System.Diagnostics.Debug.WriteLine(String.Format("DeleteEmptyFolders Skipping {0}",
						s.FilePath));

					continue;
				}

				Folder f = s.GetRootFolder() as Outlook.Folder;
				DeleteEmptyFolders(f.Folders, 0);

			}
		}
		static int DeletedFolders = 0;
		private static void DeleteEmptyFolders(Folders F, int Depth)
		{
			foreach (Folder f in F)
			{
				DeleteEmptyFolders(f.Folders, Depth + 1);
				string s = f.FullFolderPath;
				Items I = f.Items;
				int n = I.Count;
				int N = f.Folders.Count;
				if (f.DefaultItemType != OlItemType.olMailItem)
				{
					continue;
				}
				if ((n == 0) && (N == 0) && !IsPrecious(f, Depth))
				{
					Debug.WriteLine(String.Format("DELETED:{0}:{1}:{2}:{3}:{4}", DeletedFolders, Depth, n, N, s));
					f.Delete();
					DeletedFolders++;
					continue;
				}
				//Debug.WriteLine(String.Format("{0}:{1}:{2}:{3}", Depth, n, N, s));
				//Explorer p = f.Application.ActiveExplorer();
				//p.ShowPane(OlPane.olFolderList, true);
				//p.CurrentFolder = f;
			}
		}

		private static bool IsPrecious(Folder f, int Depth)
		{
			if (f.DefaultItemType != OlItemType.olMailItem)
			{
				return true;
			}
			//dynamic p = f.UserPermissions;

			if (Depth == 0)
			{

				if (f.Name.Contains("Deleted Items"))
				{
					return true;
				}
				if (f.Name.Contains("Drafts"))
				{
					return true;
				}
				if (f.Name.Contains("Sent Items"))
				{
					return true;
				}
				if (f.Name.Contains("Outbox"))
				{
					return true;
				}
				if (f.Name.Contains("Sent Items"))
				{
					return true;
				}
				if (f.Name.Contains("Junk Items"))
				{
					return true;
				}
				if (f.Name.Contains("Junk E-mail"))
				{
					return true;
				}
				if (f.Name.Contains("Junk Email"))
				{
					return true;
				}

				if (f.Name.Contains("RSS Feed"))
				{
					return true;
				}

				if (f.Name.Contains("Conversation Action Settings"))
				{
					return true;
				}
				if (f.Name.Contains("Quick Step Settings"))
				{
					return true;
				}
				if (f.Name.Contains("Inbox"))
				{
					return true;
				}
			}

			return false;
		}

		private static void CreateSortStores(Outlook.Application OA)
		{
			String PrePath = @"L:\\Outlook on L\";
			NameSpace S = OA.Session;
			for (int yrs = 1976; yrs < 2019; yrs++)
			{
				String PST_displayName = String.Format("{0}-00-00", yrs.ToString("0000"));
				String PST_name = String.Format(@"{0}\{1}.pst", PrePath, PST_displayName);
				Store St = GetStore(S, PST_displayName);
				if (St == null)
					try
					{
						S.AddStoreEx(PST_name, OlStoreType.olStoreUnicode);
					}
					catch (System.Exception e)
					{
						Debug.WriteLine(e);
						break;
					}
				Debug.WriteLine(String.Format("Store {0}", PST_displayName));
			}
		}
		private static void NameSortStores(Outlook.Application OA)
		{
			NameSpace S = OA.Session;
			for (int yrs = 2000; yrs < 2019; yrs++)
			{
				String PST_displayName = String.Format("{0}-00-00", yrs.ToString("0000"));
				Store St = GetStore(S, PST_displayName);
				if (St == null)
					continue;
				Folder rF = St.GetRootFolder() as Folder;
				rF.Name = PST_displayName;
				Outlook.PropertyAccessor pA = St.PropertyAccessor;
				dynamic P = pA.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x3001001E");

				//pA.SetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x3001001E", PST_displayName);
				rF.ShowItemCount = OlShowItemCount.olShowTotalItemCount;
				Folder a = null;
				if (OutlookFolderExists(rF.Folders, "Inbox"))
				{
					a = GetFolder(rF, "Inbox");
				}
				else
				{
					a = rF.Folders.Add("Inbox") as Outlook.Folder;
				}
				CreateMonthSubfolders(a);

			}
		}

		private static Folder GetFolder(Store S, String v)
		{
			return GetFolder(S.GetRootFolder() as Outlook.Folder, v);
		}
		private static Folder GetFolder(Folder rF, string v)
		{
			if (rF == null)
			{
				return null as Folder;
			}

			foreach (Folder f in rF.Folders)
			{
				if (f.Name.StartsWith(v) && f.Name.EndsWith(v))
				{
					return f;
				}
				Folder ff = GetFolder(f, v);
				if (ff != null)
				{
					return ff;
				}

			}
			return null as Folder;
		}

		private static Folder GetFolder(Folder rF, string v, int D)
		{
			if (rF == null)
			{
				return null as Folder;
			}
			foreach (Folder f in rF.Folders)
			{
				if (f.Name.StartsWith(v) && f.Name.EndsWith(v))
				{
					return f;
				}

				foreach (Folder FF in f.Folders)
				{
					Folder FFF = GetFolder(FF, v);
					if (FFF != null)
					{
						return FFF;
					}
				}
			}

			return null as Folder;
		}

		private static bool OutlookFolderExists(Folders folders, String name)
		{
			foreach (Folder f in folders)
			{
				if (f.Name.EndsWith(name))
				{
					return true;
				}
			}
			return false;
		}

		private static void CreateMonthSubfolders(Folder a)
		{
			a.ShowItemCount = OlShowItemCount.olShowTotalItemCount;
			var p = a.Parent as Outlook.Folder;
			String p_name = p.Name;
			for (int i = 1; i < 12; i++)
			{
				String fN = String.Format("{0}-{1}", p_name, i.ToString("00"));
				var f = a.Folders.Add(fN, OlDefaultFolders.olFolderInbox);
				f.ShowItemCount = OlShowItemCount.olShowTotalItemCount;
				Debug.WriteLine(String.Format("{0}", f.FullFolderPath));
			}
		}
		static Store GetStore(NameSpace N, String name)
		{
			foreach (Store St in N.Stores)
			{
				if (St.FilePath.Contains(name))
					return St;
			}
			return null as Store;
		}
		static void EnumerateStores(Outlook.Stores stores)
		{
			foreach (Store S in stores)
			{
				Debug.WriteLine(String.Format(@"{0}",
					S.FilePath));

				Outlook.Folder f = S.GetRootFolder() as Outlook.Folder;
				EnumerateFolder(f, 0);
			}
		}

		static void EnumerateFolder(Outlook.Folder folder, int Depth)
		{
			if (folder == null)
				return;
			Outlook.Folders childFolders = null;
			try
			{
				childFolders = folder.Folders;
			}
			catch (System.Exception e)
			{
				System.Diagnostics.Debug.WriteLine(String.Format("Liability Blabber here {0} {1} {2} {3}",
				e,
				"first argument",
				"second argument",
				"third argument",
				"unused fourth argument"));

				return;
			}

			int nChildren = childFolders.Count;
			if (nChildren == 0)
				return;
			foreach (Folder F in childFolders)
			{
				F.ShowItemCount = OlShowItemCount.olShowTotalItemCount;
				EnumerateFolder(F, Depth + 1);
				var s = F.FullFolderPath;
				var I = F.Items;
				int n = I.Count;
				int N = F.Folders.Count;
				Debug.WriteLine(String.Format("{0}:{1}:{2}:{3}", Depth, n, N, s));
			}
		}
	}
}
