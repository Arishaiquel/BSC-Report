import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Upload, Download, Loader2 } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

export default function Home() {
  const [bscFiles, setBscFiles] = useState<FileList | null>(null);
  const [bscLoading, setBscLoading] = useState(false);
  const [fameFiles, setFameFiles] = useState<FileList | null>(null);
  const [fameLoading, setFameLoading] = useState(false);
  const { toast } = useToast();

  const handleBscUpload = async () => {
    if (!bscFiles || bscFiles.length === 0) {
      toast({
        title: "No files selected",
        description: "Please select EML files or a ZIP file containing EMLs.",
        variant: "destructive",
      });
      return;
    }

    setBscLoading(true);

    try {
      const { buildExtractionWorkbook } = await import("@/lib/localExtract");
      const blob = await buildExtractionWorkbook(bscFiles);
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
      setBscLoading(false);
    }
  };

  const handleFameUpload = async () => {
    if (!fameFiles || fameFiles.length === 0) {
      toast({
        title: "No files selected",
        description: "Please select MSG files or a ZIP file containing MSGs.",
        variant: "destructive",
      });
      return;
    }

    setFameLoading(true);

    try {
      const { buildFameExtractionWorkbook } = await import("@/lib/localExtractFame");
      const blob = await buildFameExtractionWorkbook(fameFiles);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "extracted_FAME_data.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      toast({
        title: "Success",
        description: "FAME extraction completed and Excel is downloading.",
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Failed to extract FAME data";
      toast({
        title: "Error",
        description: message,
        variant: "destructive",
      });
    } finally {
      setFameLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-6 xl:p-10">
      <div className="w-full max-w-[1850px] grid gap-8 xl:grid-cols-[1.1fr_1fr_1fr]">
        <Card className="h-full rounded-none">
          <CardHeader>
            <CardTitle className="text-2xl md:text-3xl"> How to use:</CardTitle>
          </CardHeader>
          <CardContent className="space-y-8 text-base md:text-lg">
            <section className="space-y-3">
              <p className="font-semibold text-lg md:text-xl">What it does</p>
              <p className="text-muted-foreground">
                This tool reads your uploaded email files (or ZIP folder containing email files) and exports
                an Excel file. -  Upload a zip file containing the transaction files, then click "Export to Excel". Some rows will have no Buy/RSP Amounts, these are Sell, switch, RSP Amendment, etc transactions.
              </p>
            </section>

            <section className="space-y-3">
              <p className="font-semibold text-lg md:text-xl">What it extracts</p>
              <ul className="list-disc pl-6 text-muted-foreground space-y-2">
                <li>Policy Number</li>
                <li>Submission Date</li>
                <li>Buy amount and RSP Application amount </li>
                <li>Buy product type and RSP Application product type</li>
              </ul>
            </section>

            <section className="space-y-3">
              <p className="font-semibold text-lg md:text-xl">What it ignores</p>
              <ul className="list-disc pl-6 text-muted-foreground space-y-2">
                <li>Switch, Sell, Rebalance, RSP Amendment, ETF</li>
                <li>Foreign currency transactions are in a separate column</li>
              </ul>
            </section>
          </CardContent>
        </Card>

        <Card className="w-full rounded-none">
          <CardHeader>
            <CardTitle className="text-2xl md:text-3xl">iFAST</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6 text-base md:text-lg">
            <div className="rounded-lg p-10 text-center space-y-5">
              <Upload className="w-14 h-14 text-muted-foreground mx-auto" />
              <div className="space-y-2">
                <p className="text-base md:text-lg font-medium">Select your files(Save the BSC Transactions into a zip folder -{'>'} click 'Choose Files')</p>
                <p className="text-sm md:text-base text-muted-foreground">
                  Upload .eml files or a .zip folder (processed locally in your browser)
                </p>
              </div>
              <Input 
                type="file" 
                multiple 
                accept=".eml,.zip"
                className="max-w-sm mx-auto text-base"
                onChange={(e) => setBscFiles(e.target.files)}
              />
            </div>

            <Button 
              className="w-full text-lg h-14" 
              size="lg"
              onClick={handleBscUpload}
              disabled={bscLoading || !bscFiles}
            >
              {bscLoading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Processing...
                </>
              ) : (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  Export to Excel
                </>
              )}
            </Button>
          </CardContent>
        </Card>

        <Card className="w-full rounded-none">
          <CardHeader>
            <CardTitle className="text-2xl md:text-3xl">FAME</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6 text-base md:text-lg">
            <div className="rounded-lg p-10 text-center space-y-5">
              <Upload className="w-14 h-14 text-muted-foreground mx-auto" />
              <div className="space-y-2">
                <p className="text-base md:text-lg font-medium">
                  Select your files (Save FAME emails into a zip folder -{'>'} click 'Choose Files')
                </p>
                <p className="text-sm md:text-base text-muted-foreground">
                  Upload .msg files or a .zip folder (processed locally in your browser)
                </p>
              </div>
              <Input
                type="file"
                multiple
                accept=".msg,.zip"
                className="max-w-sm mx-auto text-base"
                onChange={(e) => setFameFiles(e.target.files)}
              />
            </div>

            <Button
              className="w-full text-lg h-14"
              size="lg"
              onClick={handleFameUpload}
              disabled={fameLoading || !fameFiles}
            >
              {fameLoading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Processing...
                </>
              ) : (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  Export to Excel
                </>
              )}
            </Button>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
