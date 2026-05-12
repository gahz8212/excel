import { create } from "zustand";
import term from "../locales/industry-term.json";

interface IndustryState {
  industry: keyof typeof term;
  setIndustry: (industry: keyof typeof term) => void;
  term: keyof typeof term;
}
export const useIndustryStore = create<IndustryState>((set) => ({
  industry: "construction",
  setIndustry: (industry) => set({ industry }),
  term: "finance",
}));
export const useTranslator = () => {
  const { industry } = useIndustryStore();
  const t = (key: keyof typeof term.default) => {
    return term[industry][key] || term.default[key] || key;
  };
  return{t, industry}
};
