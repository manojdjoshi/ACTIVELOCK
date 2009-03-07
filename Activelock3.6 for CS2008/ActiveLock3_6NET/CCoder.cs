using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

using System.IO;
using System.Math;
using System.Text;

public class CCoder : IDisposable
{


	#region "Atributes"
		#endregion
	private const int gc_intPragCuloare = 245;


	#region "New/Dispose"


	public CCoder()
	{
	}

	public void Dispose()
	{

		try {
		}
		catch (Exception exc) {
		}
	}

	#endregion


	#region "Private Functions"

	private Bitmap Code(Bitmap bmpPatern, string InputText)
	{
		Bitmap bmpResult = null;

		int i = 0;
		int j = 0;
		int k = 0;

		int c = 0;
		int cr = 0;
		int cg = 0;
		int cb = 0;
		string s = null;
		System.Drawing.Color clr = default(System.Drawing.Color);

		bmpResult = new Bitmap(bmpPatern.Width, bmpPatern.Height);

		for (k = 0; k <= InputText.Length - 1; k++) {
			c = Strings.Asc(InputText.Substring(k, 1));

			s = (string)c;
			s = s.PadLeft(3, '0');

			cr = (int)s.Substring(0, 1) + 1;
			cg = (int)s.Substring(1, 1) + 1;
			cb = (int)s.Substring(2, 1) + 1;
			SetR(ref i, ref j, ref bmpPatern, ref bmpResult, ref cr);
			SetG(ref i, ref j, ref bmpPatern, ref bmpResult, ref cg);
			SetB(ref i, ref j, ref bmpPatern, ref bmpResult, ref cb);
		}

		int ii = 0;
		int jj = 0;
		for (ii = i; ii <= bmpPatern.Height - 1; ii++) {
			for (jj = 0; jj <= bmpPatern.Width - 1; jj++) {
				if ((ii == i && jj >= j) || ii > i) {
					clr = bmpPatern.GetPixel(jj, ii);
					bmpResult.SetPixel(jj, ii, clr);
				}
			}
		}

		return bmpResult;
	}

	private void SetR(ref int i, ref int j, ref Bitmap bmpPatern, ref Bitmap bmpNew, ref int cr)
	{
		bool bOk = false;
		System.Drawing.Color clr = default(System.Drawing.Color);

		try {
			do {
				clr = bmpPatern.GetPixel(j, i);

				if (bmpPatern.GetPixel(j, i).R < gc_intPragCuloare) {
					bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, (clr.R + cr) % 256, clr.G, clr.B));
					bOk = true;
				}
				else {
					bmpNew.SetPixel(j, i, clr);
				}

				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
			}
			while (!bOk);
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private void SetG(ref int i, ref int j, ref Bitmap bmpPatern, ref Bitmap bmpNew, ref int cg)
	{
		bool bOk = false;
		System.Drawing.Color clr = default(System.Drawing.Color);

		try {
			do {
				clr = bmpPatern.GetPixel(j, i);

				if (clr.G < gc_intPragCuloare) {
					bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, clr.R, (clr.G + cg) % 256, clr.B));
					bOk = true;
				}
				else {
					bmpNew.SetPixel(j, i, clr);
				}

				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
			}
			while (!bOk);
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private void SetB(ref int i, ref int j, ref Bitmap bmpPatern, ref Bitmap bmpNew, ref int cb)
	{
		bool bOk = false;
		System.Drawing.Color clr = default(System.Drawing.Color);

		try {
			do {
				clr = bmpPatern.GetPixel(j, i);

				if (clr.B < gc_intPragCuloare) {
					bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, clr.R, clr.G, (clr.B + cb) % 256));
					bOk = true;
				}
				else {
					bmpNew.SetPixel(j, i, clr);
				}

				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
			}
			while (!bOk);
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private string Decode(string strPaternFile, string strMessageFile)
	{
		Bitmap bmpPatern = null;
		Bitmap bmpMessage = null;

		char c = '\0';
		int i = 0;
		int j = 0;

		try {
			StringBuilder strResult = new StringBuilder();

			bmpPatern = new Bitmap(strPaternFile);
			bmpMessage = new Bitmap(strMessageFile);

			c = Strings.GetChar(bmpPatern, bmpMessage, i, j);
			while (c != char.MinValue) {
				strResult.Append(c);
				c = Strings.GetChar(bmpPatern, bmpMessage, i, j);
			}

			bmpPatern.Dispose();
			bmpMessage.Dispose();

			return strResult.ToString();
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private char GetChar(Bitmap bmpPatern, Bitmap bmpMessage, ref int i, ref int j)
	{
		char charResult = '\0';
		string strChar = null;
		int intCrt = 0;

		try {
			if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
				return char.MinValue;
			}

			intCrt = GetR(bmpPatern, bmpMessage, ref i, ref j);
			if (intCrt < 0) {
				return char.MinValue;
			}
			strChar = (string)intCrt;

			intCrt = GetG(bmpPatern, bmpMessage, ref i, ref j);
			if (intCrt < 0) {
				return char.MinValue;
			}
			strChar += (string)intCrt;

			intCrt = GetB(bmpPatern, bmpMessage, ref i, ref j);
			if (intCrt < 0) {
				return char.MinValue;
			}
			strChar += (string)intCrt;

			charResult = Strings.Chr((int)strChar);

			return charResult;
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private int GetR(Bitmap bmpPatern, Bitmap bmpMessage, ref int i, ref int j)
	{
		int intResult = 0;
		int intPatern = 0;
		int intMessage = 0;

		try {
			if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
				return -1;
			}
			intPatern = bmpPatern.GetPixel(j, i).R;
			intMessage = bmpMessage.GetPixel(j, i).R;

			while (intPatern == intMessage) {
				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
				if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
					return -1;
				}

				intPatern = bmpPatern.GetPixel(j, i).R;

				if (intPatern < gc_intPragCuloare) {
					intMessage = bmpMessage.GetPixel(j, i).R;
				}
				else {
					intMessage = intPatern;
				}
			}

			intResult = (((intMessage + 256) - intPatern) % 256) - 1;
			if (intResult < 0) {
				intResult = 0;
			}

			j += 1;
			if (j == bmpPatern.Width) {
				j = 0;
				i += 1;
			}

			return intResult;
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private int GetG(Bitmap bmpPatern, Bitmap bmpMessage, ref int i, ref int j)
	{
		int intResult = 0;
		int intPatern = 0;
		int intMessage = 0;

		try {
			if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
				return -1;
			}
			intPatern = bmpPatern.GetPixel(j, i).G;
			intMessage = bmpMessage.GetPixel(j, i).G;

			while (intPatern == intMessage) {
				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
				if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
					return -1;
				}

				intPatern = bmpPatern.GetPixel(j, i).G;
				if (intPatern < gc_intPragCuloare) {
					intMessage = bmpMessage.GetPixel(j, i).G;
				}
				else {
					intMessage = intPatern;
				}
			}

			intResult = (((intMessage + 256) - intPatern) % 256) - 1;
			if (intResult < 0) {
				intResult = 0;
			}

			j += 1;
			if (j == bmpPatern.Width) {
				j = 0;
				i += 1;
			}

			return intResult;
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private int GetB(Bitmap bmpPatern, Bitmap bmpMessage, ref int i, ref int j)
	{
		int intResult = 0;
		int intPatern = 0;
		int intMessage = 0;

		try {
			if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
				return -1;
			}
			intPatern = bmpPatern.GetPixel(j, i).B;
			intMessage = bmpMessage.GetPixel(j, i).B;

			while (intPatern == intMessage) {
				j += 1;
				if (j == bmpPatern.Width) {
					j = 0;
					i += 1;
				}
				if (j >= bmpPatern.Width || i >= bmpPatern.Height) {
					return -1;
				}

				intPatern = bmpPatern.GetPixel(j, i).B;
				if (intPatern < gc_intPragCuloare) {
					intMessage = bmpMessage.GetPixel(j, i).B;
				}
				else {
					intMessage = intPatern;
				}
			}

			intResult = (((intMessage + 256) - intPatern) % 256) - 1;
			if (intResult < 0) {
				intResult = 0;
			}

			j += 1;
			if (j == bmpPatern.Width) {
				j = 0;
				i += 1;
			}

			return intResult;
		}
		catch (Exception exc) {
			throw exc;
		}
	}

	private string ReadFile(string strFileText)
	{
		string strResult = null;
		StreamReader srReadText = null;

		try {
			srReadText = new StreamReader(strFileText);
			strResult = srReadText.ReadToEnd();

			return strResult;
		}
		catch (Exception exc) {
			throw exc;
		}

		finally {
			try {
				srReadText.Close();
			}
			catch {
			}

		}
	}

	#endregion


	#region "Public Functions"

	public void Code(string strPaternFile, string strMessage, string strMessageFile)
	{
		Bitmap bmpPatern = null;
		Bitmap bmpMessage = null;

		try {
			bmpPatern = new Bitmap(strPaternFile);
			bmpMessage = Code(bmpPatern, strMessage);
			//ReadFile(strTextFile))
			bmpMessage.Save(strMessageFile);

		}
		catch (Exception exc) {
			throw exc;
		}

		finally {
			try {
				bmpPatern.Dispose();
			}
			catch {
			}

			try {
				bmpMessage.Dispose();
			}
			catch {
			}

		}
	}

	public void Decode(string strPaternFile, string strMessageFile, ref string strTextFile)
	{
		StreamWriter swWrite = null;

		try {
			//swWrite = New StreamWriter(strTextFile)
			//swWrite.Write(Decode(strPaternFile, strMessageFile))
			strTextFile = Decode(strPaternFile, strMessageFile);

		}
		catch (Exception exc) {
			throw exc;
		}

		finally {
			try {
				swWrite.Close();
			}
			catch {
			}

		}
	}

	#endregion


}
