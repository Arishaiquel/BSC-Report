import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Upload, Download, Loader2 } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

const FALLBACK_VERCEL_LIMIT_BYTES = 4 * 1024 * 1024;
const EXTERNAL_API_LIMIT_BYTES = 550 * 1024 * 1024;
const EXTERNAL_EXTRACT_API_URL = (import.meta.env.VITE_EXTRACT_API_URL ?? "").trim();
const EXTRACT_API_URL = EXTERNAL_EXTRACT_API_URL || "/api/extract";
const USING_VERCEL_FALLBACK_API = EXTRACT_API_URL === "/api/extract";
const ACTIVE_UPLOAD_LIMIT_BYTES = USING_VERCEL_FALLBACK_API
  ? FALLBACK_VERCEL_LIMIT_BYTES
  : EXTERNAL_API_LIMIT_BYTES;
const ACTIVE_UPLOAD_LIMIT_LABEL = USING_VERCEL_FALLBACK_API ? "4MB" : "550MB";

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

    const totalUploadSize = Array.from(files).reduce((sum, file) => sum + file.size, 0);
    if (totalUploadSize > ACTIVE_UPLOAD_LIMIT_BYTES) {
      toast({
        title: "Upload too large",
        description:
          `Total selected files exceed ${ACTIVE_UPLOAD_LIMIT_LABEL}. Please split/compress files and retry.`,
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
      const response = await fetch(EXTRACT_API_URL, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let serverMessage = "";
        const contentType = response.headers.get("content-type") || "";

        if (contentType.includes("application/json")) {
          const body = await response.json().catch(() => null);
          serverMessage =
            typeof body?.message === "string" ? body.message : JSON.stringify(body ?? {});
        } else {
          serverMessage = await response.text().catch(() => "");
        }

        if (response.status === 413) {
          throw new Error(
            `Upload too large for the current API endpoint. Reduce file size (about ${ACTIVE_UPLOAD_LIMIT_LABEL} total is safer).`,
          );
        }

        throw new Error(
          serverMessage
            ? `${serverMessage} (HTTP ${response.status})`
            : `Extraction failed (HTTP ${response.status})`,
        );
      }

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
      <Card className="w-full max-w-lg">
        <CardHeader>
          <CardTitle>BSC Report</CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="border-2 border-dashed border-muted-foreground/25 rounded-lg p-8 text-center space-y-4">
            <Upload className="w-12 h-12 text-muted-foreground mx-auto" />
            <div className="space-y-1">
              <p className="text-sm font-medium">Select your files(Save the BSC Transactions into a zip folder -{'>'} click 'Choose Files')</p>
              <p className="text-xs text-muted-foreground">
                Upload .eml files or a .zip folder (current limit: {ACTIVE_UPLOAD_LIMIT_LABEL})
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
  );
}
