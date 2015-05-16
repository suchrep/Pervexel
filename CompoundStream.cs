using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Pervexel {

	public class CompoundStream : Stream {

		BinaryReader reader;

		int sectorSize;
		int shortSectorSize;
		int shortStreamThreshold;

		List<int> sat;
		List<int> ssat;

		int sscSector;
		int sscSize;

		List<Tuple<string, int, int>> userStreams;

		public CompoundStream (Stream stream) {

			if (stream == null)
				throw new ArgumentNullException("stream");

			if (!stream.CanRead || !stream.CanSeek)
				throw new ArgumentException("stream");

			reader = new BinaryReader(stream, Encoding.Unicode);
			reader.BaseStream.Position = 0;

			var header = reader.ReadBytes(76);

			if (BitConverter.ToUInt64(header, 0) != 0xE11AB1A1E011CFD0)
				throw new Exception("THERE IS NO MAGIC");

			sectorSize = 1 << BitConverter.ToUInt16(header, 30);
			shortSectorSize = 1 << BitConverter.ToUInt16(header, 32);
			shortStreamThreshold = BitConverter.ToInt32(header, 56);

			// master sector allocation table

			var msat = new List<int>();

			for (int i = 0; i < 109; i++) {
				msat.Add(reader.ReadInt32());
			}

			for (var msatNext = BitConverter.ToInt32(header, 68); msatNext != -2; ) {
				reader.BaseStream.Position = 512 + msatNext * sectorSize;
				for (int i = 0; i < sectorSize - 4; i += 4) {
					msat.Add(reader.ReadInt32());
				}
				msatNext = reader.ReadInt32();
			}

			// sector allocation table

			sat = new List<int>();

			foreach (var satSector in msat) {
				if (satSector < 0)
					continue;
				reader.BaseStream.Position = 512 + satSector * sectorSize;
				for (int i = 0; i < sectorSize; i += 4) {
					sat.Add(reader.ReadInt32());
				}
			}

			// short-stream sector allocation table

			ssat = new List<int>();

			for (var ssatNext = BitConverter.ToInt32(header, 60); ssatNext != -2; ) {
				reader.BaseStream.Position = 512 + ssatNext * sectorSize;
				for (var i = 0; i < sectorSize; i += 4) {
					ssat.Add(reader.ReadInt32());
				}
				ssatNext = sat[ssatNext];
			}

			// directory

			userStreams = new List<Tuple<string, int, int>>();

			for (var dirNext = BitConverter.ToInt32(header, 48); dirNext != -2; ) {
				reader.BaseStream.Position = 512 + dirNext * sectorSize;
				for (var i = 0; i < sectorSize; i += 128) {
					var entry = reader.ReadBytes(128);
					// 00..64 name bytes
					// 64..66 name length
					// 66..67 type
					// 67..68 color
					// 68..72 left sibling index
					// 72..76 right sibling index
					// 76..80 first child index
					// 80..116 reserved
					// 116-120 sector
					// 120-124 size
					// 124-128 dunno
					var type = entry[66];
					if (type == 5) {
						sscSector = BitConverter.ToInt32(entry, 116);
						sscSize = BitConverter.ToInt32(entry, 120);
					} else if (type == 2) {
						var name = Encoding.Unicode.GetString(entry, 0, BitConverter.ToUInt16(entry, 64) - 2);
						var sector = BitConverter.ToInt32(entry, 116);
						var size = BitConverter.ToInt32(entry, 120);
						userStreams.Add(Tuple.Create(name, size, sector));
					}
				}
				dirNext = sat[dirNext];
			}
		}

		class Segment {
			internal int Position;
			internal int Offset;
			internal int Size;
			public override string ToString () {
				return string.Format("Position = {0}, Offset = {1}, Size = {2}", Position, Offset, Size);
			}
		}

		private IEnumerable<Segment> GetSegments (int sector, int length, bool tryShort = true) {
			if (length >= shortStreamThreshold || !tryShort) {
				var left = length;
				while (left > 0) {
					if (left < sectorSize) {
						yield return new Segment { 
							Position = length - left, 
							Offset = 512 + sector * sectorSize, 
							Size = left 
						};
						break;
					}
					yield return new Segment {
						Position = length - left,
						Offset = 512 + sector * sectorSize,
						Size = sectorSize
					};
					left -= sectorSize;
					sector = sat[sector];
				}
			} else {
				var segments = GetSegments(sscSector, sscSize, false).ToList();
				var left = length;
				while (left > 0) {
					var offset = segments[sector * shortSectorSize / sectorSize].Offset + (sector * shortSectorSize) % sectorSize;
					if (left < shortSectorSize) {
						yield return new Segment {
							Position = length - left,
							Offset = offset,
							Size = left
						};
						break;
					}
					yield return new Segment {
						Position = length - left,
						Offset = offset,
						Size = shortSectorSize
					};
					left -= shortSectorSize;
					sector = ssat[sector];
				}
			}
		}

		LinkedListNode<Segment> data;
		int position;
		int length;

		public bool Open (string name) {
			var userStream = userStreams.FirstOrDefault(s => s.Item1 == name);
			if (userStream == null) {
				data = null;
				return false;
			}
			position = 0;
			length = userStream.Item2;
			data = new LinkedList<Segment>(GetSegments(userStream.Item3, userStream.Item2)).First;
			reader.BaseStream.Position = data.Value.Offset;
			return true;
		}

		#region stream

		public override bool CanRead {
			get { return true; }
		}

		public override bool CanSeek {
			get { return true; }
		}

		public override bool CanWrite {
			get { return false; }
		}

		public override void Flush () {
			throw new NotSupportedException();
		}

		public override long Length {
			get { return length; }
		}

		public override long Position {
			get {
				return position;
			}
			set {
				Seek(value, SeekOrigin.Begin);
			}
		}

		public override int Read (byte[] buffer, int offset, int count) {
			
			var available = length - position;
			if (available == 0)
				return 0;

			var segmentOffset = position - data.Value.Position;
			var segmentLeft = data.Value.Size - segmentOffset;

			count = Math.Min(available, Math.Min(segmentLeft, count));

			var read = reader.BaseStream.Read(buffer, offset, count);

			position += read;

			if (segmentOffset + read == data.Value.Size && available > read) {
				data = data.Next;
				reader.BaseStream.Position = data.Value.Offset;
			}

			return read;
		}

		public override long Seek (long offset, SeekOrigin origin) {
			
			var targetPosition = (int)offset;
			
			if (origin == SeekOrigin.Current)
				targetPosition += position;
			else if (origin == SeekOrigin.End)
				targetPosition += length;

			if (targetPosition < 0)
				targetPosition = 0;
			else if (targetPosition > length)
				targetPosition = length;

			data = data.List.First;
			while (true) {
				if (targetPosition >= data.Value.Position && targetPosition < data.Value.Position + data.Value.Size)
					break;
				if (data.Next == null)
					break;
				data = data.Next;
			}
			
			reader.BaseStream.Position = data.Value.Offset + (targetPosition - data.Value.Position);
			position = targetPosition;
			
			return position;
		}

		public override void SetLength (long value) {
			throw new NotSupportedException();
		}

		public override void Write (byte[] buffer, int offset, int count) {
			throw new NotSupportedException();
		}

		#endregion
	}
}
