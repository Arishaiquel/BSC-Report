import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Upload, FileText, Download, Loader2 } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

export default function Home() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [loading, setLoading] = useState(false);
  const { toast } = useToast();

  const handleUpload = async () => {
    if (!files || files.length === 0) {
      toast({
        title: "No files selected",
        description: "Please select EML files or a ZIP file containing EMLs.",
        variant: "destructive",
      });
      return;
    }

    setLoading(true);
    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append("files", files[i]);
    }

    try {
      const response = await fetch("/api/extract", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) throw new Error("Extraction failed");

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "extracted_BSC_data.xlsx";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      
      toast({
        title: "Success",
        description: "Your data has been extracted and the Excel file is downloading.",
      });
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to extract data. Please check your files.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-4">
      <Card className="w-full max-w-lg">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <FileText className="w-6 h-6 text-primary" />
            BSC Report 
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="border-2 border-dashed border-muted-foreground/25 rounded-lg p-8 text-center space-y-4">
            <Upload className="w-12 h-12 text-muted-foreground mx-auto" />
            <div className="space-y-1">
              <p className="text-sm font-medium">Select your files(Save the BSC Transactions into a zip folder -{'>'} click 'Upload')</p>
              <p className="text-xs text-muted-foreground">Upload .eml files or a .zip folder</p>
            </div>
            <Input 
              type="file" 
              multiple 
              accept=".eml,.zip"
              className="max-w-xs mx-auto"
              onChange={(e) => setFiles(e.target.files)}
            />
          </div>

          <Button 
            className="w-full" 
            size="lg"
            onClick={handleUpload}
            disabled={loading || !files}
          >
            {loading ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Processing...
              </>
            ) : (
              <>
                <Download className="mr-2 h-4 w-4" />
                Extract to Excel
              </>
            )}
          </Button>
        </CardContent>
      </Card>
    </div>
  );
}
