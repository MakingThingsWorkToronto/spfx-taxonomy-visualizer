import * as React from 'react';
import styles from './TaxonomyVisualizer.module.scss';
import { ITaxonomyVisualizerProps } from './ITaxonomyVisualizerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Link, Shimmer } from 'office-ui-fabric-react';
import { Term } from '../../../services/TaxonomyService';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { IColumnBreakpoints } from '../../../models/IColumnBreakpoints';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface ITaxonomyVisualizerState {
  terms: Term[];
  loaded: boolean;
  columns:number;
  WebPartZone: HTMLElement;
}

export default class TopicsExpertise extends React.Component<ITaxonomyVisualizerProps, ITaxonomyVisualizerState> {

  private _resizeTimer: number;
  private _resizeObserver: ResizeObserver;
  private _orderedBreakpoints: IColumnBreakpoints[] = [];

  constructor(props:ITaxonomyVisualizerProps) {
    super(props);
    
    if(this.props.breakpoints && this.props.breakpoints.length > 0) {
      this._orderedBreakpoints = this.props.breakpoints.sort((a,b)=>{ return a.minPixels < b.minPixels ? -1 : 1; });
    }

    const webPartZone = this.getWebPartZone(props.domElement);
    const width : number = webPartZone.getBoundingClientRect().width;
    const columns = this.getColumnsFromWidth(width, 1);

    this.state = {
      terms: [],
      loaded: false,
      WebPartZone : webPartZone,
      columns: columns
    };

  }

  public async componentDidMount() : Promise<void> {
    
    this._resizeObserver = new ResizeObserver(this.handleResize.bind(this));
    this._resizeObserver.observe(this.state.WebPartZone);

    if(this.props.termSetId && typeof this.props.lcid !== "undefined") {

      const terms = await this.props.service.getTermSetTerms(this.props.lcid);

      this.setState({
        loaded: true,
        terms: terms
      });

    }

  }

  public componentWillUnmount() {
    this._resizeObserver.disconnect();
  }

  public render(): React.ReactElement<ITaxonomyVisualizerProps> {
    return <div className={styles.topicsExpertiseWebPart}>
        <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateProperty}
              themeVariant={this.props.theme}
              className={styles.webPartTitle}
              />
          {this.state.loaded ? this.renderTopics() : this.renderShimmer()}
      </div>;
  }

  private renderShimmer() : JSX.Element {
    const columnClassName : string = this.getColumns();
    let shimmers : JSX.Element[] = [];
    for(var i = 0; i<this.state.columns; i++) {  shimmers.push(this.renderShimmerBlock()); }

    return (
      <div className={ styles.topicsExpertise + " " + columnClassName }>
        {shimmers}
      </div>
    );
  }

  private renderShimmerBlock() : JSX.Element {
    return (<div className={styles.shimmerBlock}>
      <Shimmer className={styles.shimmerLine}></Shimmer>
      <Shimmer className={styles.shimmerLine}></Shimmer>
      <Shimmer className={styles.shimmerLine}></Shimmer>
    </div>);
  }

  private renderTopics() : JSX.Element {
    const extraColumns = (this.state.terms || []).length % this.state.columns;
    const placeHolderElements : JSX.Element[] = [];
    for(var i=0;i<extraColumns;i++){
      placeHolderElements.push(<div className={styles.topicGroup}>&nbsp;</div>);
    }
    const columnClassName : string = this.getColumns();
    return (
      <div className={ styles.topicsExpertise + " " + columnClassName }>
        {this.state.terms.map((term: Term)=>{
          return this.renderLinkGroup(term);
        })}
        {placeHolderElements}
      </div>
    );
  }

  private getColumns():string {
    if(typeof this.state.columns === "undefined") return styles.col1;
    const styleName = "col" + this.state.columns.toString();
    return styles[styleName];
  }

  private renderLinkGroup(term: Term) : JSX.Element {
    return (
      <div className={styles.topicGroup}>
        <div className={styles.topicHeader}>{this.renderHeaderLink(term)}</div>
        {term.Children.map((childTerm:Term)=>{
          return <div className={styles.topicChild}>{this.renderLink(childTerm)}</div>;
        })}
      </div>
    );
  }

  private renderHeaderLink(term: Term) : JSX.Element {
    const url = this.getLinkHref(term);
    return <Link href={url} style={{color:this.props.theme.semanticColors.bodyText}}>{term.Label}</Link>;
  }

  private renderLink(term: Term) : JSX.Element {
    const url = this.getLinkHref(term);
    return <Link href={url} style={{color:this.props.theme.semanticColors.link}}>{term.Label}</Link>;
  }

  private getLinkHref(term:Term) : string {
    
    return !this.props.linkTemplate 
      ? ""
      : this.props.linkTemplate
          .replace(/\{TermLabel\}/g, term.Label)
          .replace(/\{TermGuid\}/g, term.Id)
          .replace(/\{TermSetId\}/g, this.props.termSetId)
          .replace(/\{TermName\}/g, term.Name);

  }

  private resizeEvent(entries:any) {
    const webPart = entries[0];

    if(webPart) {
      
      const width = webPart.contentRect.width;

      this.setState({
        columns: this.getColumnsFromWidth(width, this.state.columns)
      });

    }
  }

  private handleResize(entries:any) {
    window.clearTimeout(this._resizeTimer);
    this._resizeTimer = window.setTimeout(this.resizeEvent.bind(this,entries), 500);
  }

  private getColumnsFromWidth(width: number, cols: number) : number {
    let columns = cols;
    this._orderedBreakpoints.map((breakpoint)=> {
      if(breakpoint.minPixels <= width) {
        columns = breakpoint.columns;
      }
    });
    return columns;
  }
  
  private getWebPartZone(element:HTMLElement) : HTMLElement {
    return this.props.displayMode === DisplayMode.Edit
      ? element.closest(".ControlZone--edit") as HTMLElement
      : (element.closest(".ControlZone--control") as HTMLElement) || (element.closest(".ControlZone-control") as HTMLElement) || element.closest(".ControlZone") as HTMLElement;
      
  }

}
