namespace WindowsFormsApplication1
{
    internal class OpenFileDialog
    {
        public bool Multiselect { get; internal set; }
        public string DefaultExt { get; internal set; }
        public string Filter { get; internal set; }
        public string Title { get; internal set; }
        public string FileName { get; internal set; }

        internal object ShowDialog()
        {
            throw new NotImplementedException();
        }
    }
}