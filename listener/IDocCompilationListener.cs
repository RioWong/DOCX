
namespace CreateWord.listener
{
    public enum DocCompilationStatus { Success, Fail }
    public class DocCompilationArg
    {
        public int companyId;
        public string docPath;
        public DocCompilationStatus docCompilationStatus;
        public string reasonOfFailed;

        public DocCompilationArg(int companyId, string docPath, DocCompilationStatus docCompilationStatus)
        {
            this.companyId = companyId;
            this.docPath = docPath;
            this.docCompilationStatus = docCompilationStatus;
            this.reasonOfFailed = "";
        }

        public DocCompilationArg(int companyId, string docPath,
            DocCompilationStatus docCompilationStatus, string reasonOfFailed)
        {
            this.companyId = companyId;
            this.docPath = docPath;
            this.docCompilationStatus = docCompilationStatus;
            this.reasonOfFailed = reasonOfFailed;
        }
    }

    public interface IDocCompilationListener
    {
        void DocCompleted( DocCompilationArg arg );
    }
}
