import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { slideLibraryApi } from "./services/api";
import { createPowerPointService } from "./services/powerpoint";

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("The application root element was not found.");
}

const root = createRoot(rootElement);

const renderApplication = async (): Promise<void> => {
  const powerPointService = await createPowerPointService(slideLibraryApi);

  root.render(
    <StrictMode>
      <App api={slideLibraryApi} powerPointService={powerPointService} />
    </StrictMode>
  );
};

void renderApplication();
