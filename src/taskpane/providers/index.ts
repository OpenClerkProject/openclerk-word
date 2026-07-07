import { citationProviderRegistry } from "./registry";
import { CourtListenerProvider } from "./courtListenerProvider";
import { LexisNexisProvider } from "./lexisNexisProvider";
import { WestlawProvider } from "./westlawProvider";
import { BloombergLawProvider } from "./bloombergLawProvider";
import { OpenCaseProvider } from "./openCaseProvider";
import { UsptoPatentCenterProvider } from "./usptoPatentCenterProvider";

citationProviderRegistry.register(new CourtListenerProvider());
citationProviderRegistry.register(new LexisNexisProvider());
citationProviderRegistry.register(new WestlawProvider());
citationProviderRegistry.register(new BloombergLawProvider());
citationProviderRegistry.register(new OpenCaseProvider());
citationProviderRegistry.register(new UsptoPatentCenterProvider());

export { citationProviderRegistry } from "./registry";
export * from "./types";
export { parseCaseCitation, extractCaseCitations } from "./citationParser";
