using System.Collections.Generic;

namespace Delaney.DocX
{
    public interface IBlockLevelContent : IElement
    {
        List<IMedia> Medias { get; }
    }
}