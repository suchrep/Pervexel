using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Pervexel {

	public class WorkbookReader {

		BinaryReader reader;
		string[] sst;
		DateTime baseDate;
		Dictionary<int, string> formats;
		List<int> formatMap;
		List<Tuple<string, uint>> sheets;

		IEnumerator<object> enumerator;
		int row;
		int col;
		object val;

		public int Row { get { return row; } }
		public int Column { get { return col; } }
		public object Value { get { return val; } }

		public WorkbookReader (Stream stream) {

			if (stream == null)
				throw new ArgumentNullException("stream");

			if (!stream.CanRead || !stream.CanSeek)
				throw new ArgumentException("stream");

			this.reader = new BinaryReader(stream);

			var record = ReadRecord();
			if (record.Id != 0x0809 && BitConverter.ToUInt16(record.Data, 2) != 5) {
				throw new Exception("EXPECTO GLOBALS");
			}

			sheets = new List<Tuple<string, uint>>();
			formats = new Dictionary<int, string>();
			formatMap = new List<int>();
			baseDate = new DateTime(1900, 1, 1);

			while (true) {
				
				record = ReadRecord();
				var id = record.Id;
				var data = record.Data;

				// WORKSHEET
				if (id == 0x0085) {
					var streamOffset = BitConverter.ToUInt32(data, 0);
					var state = data[4];					
					var type = data[5];
					var len = data[6];
					var options = data[7];
					var name = (options & 1) != 0 ? Encoding.Unicode.GetString(data, 8, len * 2) : new string(data.Skip(8).Take(len).Select(b => (char)b).ToArray());
					// 0 - sheet, 2 - chart, 6 - vb
					if (type == 0) {
						sheets.Add(Tuple.Create(name, streamOffset));
					}
				}

				// SST
				if (id == 0x00FC) {
					var count = BitConverter.ToInt32(data, 4);
					sst = new string[count];
					var offset = 8;
					for (var i = 0; i < count; i++) {
						var value = GetString(data, ref offset);
						while (offset < 0 || offset == data.Length && i + 1 < count) {
							var next = ReadRecord();
							if (next == null || next.Id != 0x003C)
								throw new Exception("EXPECTO CONTINUE");
							id = next.Id;
							data = next.Data;
							if (offset < 0) {
								var nextPart = GetString(data, ref offset);
								value += nextPart;
							} else {
								offset = 0;
								break;
							}
						}
						sst[i] = value;
					}
				}

				// FORMAT
				if (id == 0x041E) {
					var formatIndex = BitConverter.ToUInt16(data, 0);
					int offset = 2;
					var formatString = GetString(data, ref offset);
					formats[formatIndex] = formatString;
				}

				// EXTENDED FORMAT
				if (id == 0x00E0) {
					var formatIndex = BitConverter.ToUInt16(data, 2);
					formatMap.Add(formatIndex);
				}

				// DATEMODE
				if (id == 0x0022) {
					if (data[0] == 1) {
						baseDate = new DateTime(1904, 1, 2);
					}
				}

				if (id == 0x000A)
					break;
			}
		}

		class Record {
			public int Id;
			public byte[] Data;
		}

		Record ReadRecord () {
			var id = reader.ReadUInt16();
			var size = reader.ReadUInt16();
			var bytes = reader.ReadBytes(size);
			return new Record { Id = id, Data = bytes };
		}

		public bool Open (int index) {
			if (index < 0 || index > sheets.Count)
				return false;
			reader.BaseStream.Position = sheets[index].Item2;
			enumerator = GetEnumerator();
			return true;
		}

		public bool Read () {
			if (enumerator.MoveNext()) {
				val = enumerator.Current;
				return true;
			}
			return false;
		}

		private IEnumerator<object> GetEnumerator () {

			while (true) {

				var record = ReadRecord();
				var id = record.Id;
				var data = record.Data;

				// LABEL
				if (id == 0x0207) {
					var offset = 0;
					var result = GetString(data, ref offset);
					yield return result;
				}

				// LABELSST
				if (id == 0x00FD) {
					row = BitConverter.ToUInt16(data, 0);
					col = BitConverter.ToUInt16(data, 2);
					var sstIndex = BitConverter.ToInt32(data, 6);
					var value = sst[sstIndex];
					yield return value;
				}

				// NUMBER
				if (id == 0x0203) {
					row = BitConverter.ToUInt16(data, 0);
					col = BitConverter.ToUInt16(data, 2);
					var xf = BitConverter.ToUInt16(data, 4);
					var value = BitConverter.ToDouble(data, 6);
					var result = Format(value, xf);
					yield return result;
				}

				// RK
				if (id == 0x027E) {
					row = BitConverter.ToUInt16(data, 0);
					col = BitConverter.ToUInt16(data, 2);
					var xf = BitConverter.ToUInt16(data, 4);
					var value = GetRK(data, 6);
					var result = Format(value, xf);
					yield return result;
				}

				// MULRK
				if (id == 0x00BD) {
					row = BitConverter.ToUInt16(data, 0);
					col = BitConverter.ToUInt16(data, 2);
					var offset = 4;
					while (data.Length - offset >= 6) {
						var xf = BitConverter.ToUInt16(data, offset);
						offset += 2;
						var value = GetRK(data, offset);
						var result = Format(value, xf);
						yield return result;
						col++;
						offset += 4;
					}
				}

				// FORMULA
				if (id == 0x0006) {
					row = BitConverter.ToUInt16(data, 0);
					col = BitConverter.ToUInt16(data, 2);
					var xf = BitConverter.ToUInt16(data, 4);
					if (BitConverter.ToUInt16(data, 6 + 6) == 0xFFFF) {
						if (data[6] == 0) {
							// string follows
						} else if (data[6] == 1) {
							var value = (data[6 + 2] != 0);
							yield return value;
						} else if (data[6] == 2) {
							// error
						} else if (data[6] == 3) {
							// empty
						}
					} else {
						var value = BitConverter.ToDouble(data, 6);
						var result = Format(value, xf);
						yield return result;
					}
				}

				if (id == 0x000A)
					break;
			}
		}

		object Format (double value, int xf) {
			var format = formatMap[xf];
			// these are dates by design, custom date formats aren't supported
			if (format >= 14 && format <= 17) {
				var result = baseDate.AddDays(value);
				return result;
			}
			return value;
		}

		static double GetRK (byte[] bytes, int offset) {

			long raw = BitConverter.ToInt32(bytes, offset);
			var div100 = (raw & 1) != 0;
			var isInt = (raw & 2) != 0;

			raw = raw >> 2;

			var result = isInt ? (double)raw : BitConverter.Int64BitsToDouble(raw << 34);

			if (div100)
				result /= 100;

			return result;
		}

		static string GetString (byte[] bytes, ref int offset) {

			ushort len;
			if (offset >= 0) {
				len = BitConverter.ToUInt16(bytes, offset);
				offset += 2;
			} else {
				len = (ushort)-offset;
				offset = 0;
			}

			var options = bytes[offset];
			offset += 1;

			var compressed = (options & 0x01) == 0;

			int charsToRead = bytes.Length - offset;
			if (!compressed)
				charsToRead /= 2;

			charsToRead = Math.Min(len, charsToRead);

			string result;
			
			if (compressed) {
				var chars = new char[charsToRead];
				Array.Copy(bytes, offset, chars, 0, charsToRead);
				result = new string(chars);
				offset += len;
			} else {
				result = Encoding.Unicode.GetString(bytes, offset, charsToRead * 2);
				offset += len * 2;
			}

			if (charsToRead < len)
				offset = charsToRead - len;

			return result;
		}
	}
}
