using System.IO.Compression;

namespace Delaney.DocX
{
    public interface IMedia : IElement
    {
        int Id { get; set; }
        string RelationshipId { get; set; }

        string Name { get; set; }

        string SourceFullname { get; set; }
        string TargetFullname { get; set; }
        string Target { get; set; }
        string FileExtension { get; }
        string FilenameWithoutExtension { get;  }

        void CopyTo(ZipArchive zipArchive_Target,
                    string fullname_Target);
    }
}