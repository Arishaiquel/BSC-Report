import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Upload, Download, Loader2 } from "lucide-react";
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

    try {
      const { buildExtractionWorkbook } = await import("@/lib/localExtract");
      const blob = await buildExtractionWorkbook(files);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "extracted_BSC_data.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      
      toast({
        title: "Success",
        description: "Extraction completed in your browser and Excel is downloading.",
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Failed to extract data";
      toast({
        title: "Error",
        description: message,
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-4">
      <div className="w-full max-w-6xl grid gap-6 lg:grid-cols-[1.2fr_1fr]">
        <Card className="h-full">
          <CardHeader>
            <CardTitle> How to use:</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6 text-sm">
            <section className="space-y-2">
              <p className="font-semibold">What it does</p>
              <p className="text-muted-foreground">
                This tool reads your uploaded email files (or ZIP folder containing email files) and exports
                an Excel for the BSC transaction report. - you can just upload the downloaded ZIP folder from iFAST. Some rows will have no Buy/RSP Amounts, this are Sell, switch, RSP Amendment, etc transactions.
              </p>
            </section>

            <section className="space-y-2">
              <p className="font-semibold">What it extracts</p>
              <ul className="list-disc pl-5 text-muted-foreground space-y-1">
                <li>Policy Number</li>
                <li>Submission Date</li>
                <li>Buy amount and RSP Application amount </li>
                <li>Buy product type and RSP Application product type</li>
              </ul>
            </section>

            <section className="space-y-2">
              <p className="font-semibold">What it ignores</p>
              <ul className="list-disc pl-5 text-muted-foreground space-y-1">
                <li>Switch, Sell, Rebalance, RSP Amendment, ETF</li>
                <li>Foreign currency transactions are in a separate column</li>
                <li>Files that are not .eml (except .zip files)</li>
              </ul>
            </section>
          </CardContent>
        </Card>

        <Card className="w-full">
          <CardHeader>
            <CardTitle>BSC Report</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="border-2 border-dashed border-muted-foreground/25 rounded-lg p-8 text-center space-y-4">
              <Upload className="w-12 h-12 text-muted-foreground mx-auto" />
              <div className="space-y-1">
                <p className="text-sm font-medium">Select your files(Save the BSC Transactions into a zip folder -{'>'} click 'Choose Files')</p>
                <p className="text-xs text-muted-foreground">
                  Upload .eml files or a .zip folder (processed locally in your browser)
                </p>
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
    </div>
  );
}
