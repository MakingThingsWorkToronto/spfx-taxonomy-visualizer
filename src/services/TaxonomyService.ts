import { taxonomy, ITermSet, ITerm, ITerms } from "@pnp/sp-taxonomy";
import { sp } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface Label {
    IsDefaultForLanguage: boolean;
    Language: number;
    Value: string;
}

export interface Term {
    Children: Term[];
    Parent:string;
    TermCount:number;
    Name:string;
    AllLabels:Label[];
    Label: string;
    Id: string;
    [key:string]:any;
}

export class TaxonomyService {
    
    private _termSetId: string = "";

    constructor(context:WebPartContext, termSetId: string){
        sp.setup({
            spfxContext: context
        });
        this._termSetId = termSetId;
    }

    public async getTermSetTerms(lcid: number) : Promise<Term[]> {

        if(!this._termSetId || this._termSetId.trim().length === 0) return [];

        const termSet : ITermSet = taxonomy.getDefaultKeywordTermStore().getTermSetById(this._termSetId);
        
        return await this.getTerms(termSet.terms, lcid);
        
    }

    private async getTerms(terms:ITerms, lcid: number) : Promise<Term[]> {

        const all: (ITerm)[] = await terms.select("Name", "Id", "TermsCount", "Parent", "Labels", "CustomProperties", "LocalCustomProperties").get();
        let ret:Term[] = [];
        let termHash:object = {};
        
        for(let i = 0; i<all.length; i++) {

            const data = all[i] as any;
            const termId = this.parseGuid(data.Id);
            const children : Term[] = [];
            const name = data.Name;
            const labels = this.unpackLabels(data.Labels._Child_Items_);
            const parentId = !data.Parent ? null : this.parseGuid(data.Parent.Id);
            let defaultLabel : string = name;
            
            labels.some(label=> {
                if(label.Language === lcid && label.IsDefaultForLanguage) {
                    defaultLabel = label.Value;
                    return true;
                }
            });

            let newTerm : Term = {
                Id : termId,
                Children: children,
                Parent: parentId,
                TermCount: data.TermsCount,
                Name: name,
                AllLabels : labels,
                Label: defaultLabel,
                ...data.CustomProperties,
                ...data.LocalCustomProperties
            };

            if(!data.Parent) {
                ret.push(newTerm);
            } else {
                if(!termHash[parentId]) termHash[parentId] = [];
                termHash[parentId].push(newTerm);
            }

        }

        return this.makeHierarchy(ret, termHash);

    }

    private parseGuid(test:string):string {
        if(!test) return "";
        return test.replace("/Guid(","").replace(")/","");
    }

    private makeHierarchy(parents:Term[], termHash:object) : Term[] {
        let ret:Term[] = [];
        parents.map((parent)=> {
            if(parent.TermCount > 0 && termHash[parent.Id] && termHash[parent.Id].length > 0) {
                parent.Children = this.makeHierarchy(termHash[parent.Id], termHash);   
            }
            ret.push(parent);
        });

        return ret;
    }

    private unpackLabels(labels:any[]) : Label[] {
        return labels.map((label:any)=> {
            return {
                IsDefaultForLanguage: label.IsDefaultForLanguage,
                Language: label.Language,
                Value: label.Value
            };
        });
    }



}