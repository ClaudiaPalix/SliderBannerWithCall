import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SliderBannerWithCallWebPart.module.scss';
import * as strings from 'SliderBannerWithCallWebPartStrings';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISliderBannerWithCallWebPartProps {
  description: string;
  BannerImageUrl: {
    Url: string;
  };
  Title: string;
  FileLeafRef: string;
  image: string;
}

export default class SliderBannerWithCallWebPart extends BaseClientSideWebPart<ISliderBannerWithCallWebPartProps> {
  
  private slideIndex: number = 1;
  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value;
        console.log('User Email using User Principal Name:', this.userEmail);
        // Now you can use this.userEmail as needed
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  } 

  public render(): void {
    const decodedDescription = decodeURIComponent(this.properties.description); // Decode the description (like incase there is blank space, or special characters, etc)
    // console.log(decodedDescription);
    // console.log("See All Button Url: ", decodedDescription);
    this.userDetails().then(() => {
      // console.log("Start of HTML");
      this.domElement.innerHTML = `
        <section class="${styles.bannerSection}">
          <div class="${styles.containerFluid}">
          <div class="${styles.rowFlex}">
              <div class="${styles.colMd6} ${styles.ParentHeight} ${styles.PositionRelative}">
                <div class="${styles.seeAllText}">
                    <a href="${decodedDescription}" target="_self">See All</a>
                </div>
                <div class="${styles.slideshowContainer}" id="slideshowContainer">
                    <div id="slides" class="${styles.slidesDiv}"></div>
                    <a class="${styles.prev}" onclick="this.plusSlides(-1)">❮</a>
                    <a class="${styles.next}" onclick="this.plusSlides(1)">❯</a>

                    <div class="${styles.dotsDiv}">
                      <span class="${styles.dot}" data-slide-index="1" onclick="this.currentSlide(1)"></span>
                      <span class="${styles.dot}" data-slide-index="2" onclick="this.currentSlide(2)"></span>
                      <span class="${styles.dot}" data-slide-index="3" onclick="this.currentSlide(3)"></span>
                      <span class="${styles.dot}" data-slide-index="4" onclick="this.currentSlide(4)"></span>
                    </div>
                </div>
              </div>
              <div class="${styles.colMd6} ${styles.ParentHeight}"">
                <div class="${styles.rowFlex} ${styles.ParentHeight}"">

                    <div class="${styles.colSm6}" id="Block1div">
                      <a id="Block1Link">
                      <div class="${styles.imgBox}"> 
                          <img id="Block1Img">
                          <div class="${styles.innerContents}">
                          <h4 id="Block1Title"></h4>
                        </div>
                      </div>
                      </a>
                    </div>

                    <div class="${styles.colSm6}" id="Block2div">
                      <a id="Block2Link">
                      <div class="${styles.imgBox}"> 
                          <img id="Block2Img">
                          <div class="${styles.innerContents}">
                          <h4 id="Block2Title"></h4>
                        </div>
                      </div>
                      </a>
                    </div>

                    <div class="${styles.colSm6}" id="Block3div">
                      <a id="Block3Link">
                      <div class="${styles.imgBox}"> 
                          <img id="Block3Img">
                          <div class="${styles.innerContents}">
                          <h4 id="Block3Title"></h4>
                        </div>
                      </div>
                      </a>
                    </div>

                    <div class="${styles.colSm6}" id="Block4div">
                      <a id="Block4Link">
                      <div class="${styles.imgBox}"> 
                          <img id="Block4Img">
                          <div class="${styles.innerContents}">
                          <h4 id="Block4Title"></h4>
                        </div>
                      </div>
                      </a>
                    </div>
                </div>
              </div>
          </div>
        </div>
    </section>`;
  //  console.log("End of HTML");
  //  console.log("Generated class names:", styles);
   this.setupEventHandlers();
  });
  }

  private _renderBlock1(): void {
    let apiUrl: string = ``;
    if(this.userEmail.includes(".admin@")){
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._checkIfAdmin()}%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block1%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }else{
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block1%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }
    
  
    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
      .then(response => response.json())
      .then(data => {
        console.log("Api response: ", data);
        if (data.value && data.value.length > 0) {
          data.value.forEach((item: any) => {

            const img: HTMLImageElement | null = this.domElement.querySelector('#Block1Img');
            if (img) {
              img.src = this._getBannerImage(item.BannerImageUrl.Url);
            } else {
              console.error("Image element not found.");
            }
            
            const a: HTMLAnchorElement | null = this.domElement.querySelector('#Block1Link');
            if (a) {
               a.onclick = () => {
              window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
              };
            } else {
              console.error("Anchor element not found.");
            }
             

            const h2: HTMLHeadingElement | null = this.domElement.querySelector('#Block1Title');
            
            if (h2) {
              h2.textContent = item.Title;
            } else {
              console.error("H2 element not found.");
            }
          });
        } else {
          const noDataMessage = 'Error for block1.';
          console.log(noDataMessage);
        }
      })
      .catch(error => {
        console.error("Error fetching user data: ", error);
      });
}

private _renderBlock2(): void {
  let apiUrl: string = ``;
  if(this.userEmail.includes(".admin@")){
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._checkIfAdmin()}%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block2%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
  }else{
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block2%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
  }
  

  fetch(apiUrl, {
    method: 'GET',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    }
  })
    .then(response => response.json())
    .then(data => {
      console.log("Api response: ", data);
      if (data.value && data.value.length > 0) {
        data.value.forEach((item: any) => {

          const img: HTMLImageElement | null = this.domElement.querySelector('#Block2Img');
          if (img) {
            img.src = this._getBannerImage(item.BannerImageUrl.Url);
          } else {
            console.error("Image element not found.");
          }
          
          const a: HTMLAnchorElement | null = this.domElement.querySelector('#Block2Link');
          if (a) {
             a.onclick = () => {
            window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
            };
          } else {
            console.error("Anchor element not found.");
          }
           

          const h2: HTMLHeadingElement | null = this.domElement.querySelector('#Block2Title');
          
          if (h2) {
            h2.textContent = item.Title;
          } else {
            console.error("H2 element not found.");
          }
        });
      } else {
        const noDataMessage = 'Error for block2.';
        console.log(noDataMessage);
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
}

private _renderBlock3(): void {
  let apiUrl: string = ``;
  if(this.userEmail.includes(".admin@")){
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._checkIfAdmin()}%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block3%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
  }else{
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block3%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
  }
  

  fetch(apiUrl, {
    method: 'GET',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    }
  })
    .then(response => response.json())
    .then(data => {
      console.log("Api response: ", data);
      if (data.value && data.value.length > 0) {
        data.value.forEach((item: any) => {

          const img: HTMLImageElement | null = this.domElement.querySelector('#Block3Img');
          if (img) {
            img.src = this._getBannerImage(item.BannerImageUrl.Url);
          } else {
            console.error("Image element not found.");
          }
          
          const a: HTMLAnchorElement | null = this.domElement.querySelector('#Block3Link');
          if (a) {
             a.onclick = () => {
            window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
            };
          } else {
            console.error("Anchor element not found.");
          }
           

          const h2: HTMLHeadingElement | null = this.domElement.querySelector('#Block3Title');
          
          if (h2) {
            h2.textContent = item.Title;
          } else {
            console.error("H2 element not found.");
          }
        });
      } else {
        const noDataMessage = 'Error for block3.';
        console.log(noDataMessage);
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
}

private _renderBlock4(): void {
  let apiUrl: string = ``;
    if(this.userEmail.includes(".admin@")){
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._checkIfAdmin()}%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block4%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }else{
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27Block4%27&$orderby=Modified%20desc&$top=1&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }
    

  fetch(apiUrl, {
    method: 'GET',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    }
  })
    .then(response => response.json())
    .then(data => {
      console.log("Api response: ", data);
      if (data.value && data.value.length > 0) {
        data.value.forEach((item: any) => {

          const img: HTMLImageElement | null = this.domElement.querySelector('#Block4Img');
          if (img) {
            img.src = this._getBannerImage(item.BannerImageUrl.Url);
          } else {
            console.error("Image element not found.");
          }
          
          const a: HTMLAnchorElement | null = this.domElement.querySelector('#Block4Link');
          if (a) {
             a.onclick = () => {
            window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
            };
          } else {
            console.error("Anchor element not found.");
          }
           

          const h2: HTMLHeadingElement | null = this.domElement.querySelector('#Block4Title');
          
          if (h2) {
            h2.textContent = item.Title;
          } else {
            console.error("H2 element not found.");
          }
        });
      } else {
        const noDataMessage = 'Error for block4.';
        console.log(noDataMessage);
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
}


  private _renderSlides(): void {
    const OuterDiv: HTMLElement | null = this.domElement.querySelector(`#slides`);
    let apiUrl: string = ``;
    if(this.userEmail.includes(".admin@")){
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._checkIfAdmin()}%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27BannerCarousel%27&$orderby=Modified%20desc&$top=4&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }else{
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20and(OData__ModernAudienceTargetUserFieldId%20eq%20%2716%27%20or%20OData__ModernAudienceTargetUserFieldId%20eq%20%27${this._getCompanyFromEmail()}%27)%20and%20ContentType0%20eq%20%27BannerCarousel%27&$orderby=Modified%20desc&$top=4&$select=Title,BannerImageUrl,FileLeafRef,ContentType0,OData__ModernAudienceTargetUserFieldId`;
    }
    
    
    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
      .then(response => response.json())
      .then(data => {
        console.log("Api response: ", data);
        if (data.value && data.value.length > 0) {
          data.value.forEach((item: any) => {
            const innerDiv: HTMLDivElement = document.createElement('div');
            innerDiv!.classList.add(styles.bannerSlides, styles.fade);
  
            const imageDiv: HTMLDivElement = document.createElement('div');
            imageDiv!.classList.add(styles.slideImageDiv);
  
            const img: HTMLImageElement = document.createElement('img');
            img.src = this._getBannerImage(item.BannerImageUrl.Url);
            imageDiv.appendChild(img);
            // console.log("Image Url for",item.Title, this._getBannerImage(item.BannerImageUrl.Url));
  
            const textDiv: HTMLDivElement = document.createElement('div');
            textDiv.classList.add(styles.text);
            textDiv.onclick = () => {
              window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
            };
            textDiv.textContent = item.Title;

            const gradientDiv: HTMLDivElement = document.createElement('div');
            gradientDiv!.classList.add(styles.bottomGradient);
            gradientDiv.onclick = () => {
              window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${item.FileLeafRef}`, '_self');
            };
  
            innerDiv!.appendChild(imageDiv);
            innerDiv!.appendChild(textDiv);
            innerDiv!.appendChild(gradientDiv);

            OuterDiv!.appendChild(innerDiv);
          });
  
          this.showSlides(this.slideIndex);
        } else {
          const noDataMessage = 'No applications available for the user.';
          console.log(noDataMessage);
        }
      })
      .catch(error => {
        console.error("Error fetching user data: ", error);
      });
  }
  
  private setupEventHandlers(): void {
    
    this._renderBlock1();
    this._renderBlock2();
    this._renderBlock3();
    this._renderBlock4();
    this._renderSlides();
    // console.log("Start of setupEventHandlers");
  
    // Defer the execution of showSlides to ensure that the DOM is fully rendered
    setTimeout(() => {
      this.showSlides(this.slideIndex);
    }, 0);
  
    document.addEventListener('click', (event: Event) => {
      const target = event.target as HTMLElement;
      if (target.classList.contains(`${styles.prev}`)) {
        this.plusSlides(-1);
      } else if (target.classList.contains(`${styles.next}`)) {
        this.plusSlides(1);
      } else if (target.classList.contains(`${styles.dot}`)) {
        const index = parseInt(target.getAttribute('data-slide-index') || '1', 10);
        this.currentSlide(index);
      }
    });
  
    setInterval(() => {
      this.plusSlides(1);
    }, 8000);
    // console.log("Slide Index: ",this.slideIndex);

    // console.log("End of setupEventHandlers");
  }

  private plusSlides(n: number): void {
    // console.log("Start of plusSlides");
    this.showSlides(this.slideIndex + n);
    // console.log("End of plusSlides");
  }

  private currentSlide(n: number): void {
    // console.log("Start of currentSlide");
    this.showSlides(n);
    // console.log("End of currentSlide");
  }

  private showSlides(n: number): void {
    // console.log("Start of showSlides");

    const slides = this.domElement.getElementsByClassName(`${styles.bannerSlides}`) as HTMLCollectionOf<HTMLElement>;
    const dots = this.domElement.getElementsByClassName(`${styles.dot}`) as HTMLCollectionOf<HTMLElement>;

    // console.log("Slides:", slides);
    // console.log("Dots:", dots);

    if (!slides || slides.length === 0) {
      console.error("No slides found");
      return;
    }

    // Adjusting index calculation
    if (n > slides.length) {
      this.slideIndex = 1;
    } else if (n < 1) {
      this.slideIndex = slides.length;
    } else {
      this.slideIndex = n;
    }

    for (let i = 0; i < slides.length; i++) {
      slides[i].style.display = "none";
    }

    for (let i = 0; i < dots.length; i++) {
      dots[i].classList.remove(`${styles.active}`);
    }

    slides[this.slideIndex - 1].style.display = "block";
    dots[this.slideIndex - 1].classList.add(`${styles.active}`);

    // console.log("End of showSlides");
}

private _checkIfAdmin(): string{
    let adminGroup: string = "";

    if(this.userEmail.includes(".admin@")){
      console.log("This user is an admin");
      adminGroup = "19";
    }else{
      console.log("User is not Admin");
    }
    return adminGroup;

}

private _getCompanyFromEmail(): string {
    let userGroup: string = "";

    if(this.userEmail.includes("@aciesinnovations.com"))
    {
      console.log("User belongs to acies");
      userGroup = "7";
    }else if(this.userEmail.includes("_zensar.com") || this.userEmail.includes("@zensar.")){
      console.log("User belongs to zensar");
      userGroup = "20";
    }else if(this.userEmail.includes("@rpg.com") || this.userEmail.includes("@rpg.in")){
      console.log("User belongs to rpg");
      userGroup = "17";
    }else if(this.userEmail.includes("@ceat.com")){
      console.log("User belongs to ceat");
      userGroup = "36";
    }else if(this.userEmail.includes("_harrisonsmalayalam.com") || this.userEmail.includes("@harrisonsmalayalam.com")){
      console.log("User belongs to harrison");
      userGroup = "18";
    }else if(this.userEmail.toLowerCase().includes("@kecrpg.com")){
      console.log("User belongs to kec");
      userGroup = "39";
    }else if(this.userEmail.includes("@raychemrpg.com")){
      console.log("User belongs to raychem");
      userGroup = "37";
    }else if(this.userEmail.includes("@rpgls.com")){
      console.log("User belongs to rpgls");
      userGroup = "38";
    }
  return userGroup;
}  

  private _getBannerImage(image: string): string {
    // console.log("Setting size for image:", image);
    let imageUrl: string = "";

    if (image.includes("hubblecontent")) {
      imageUrl = image.replace("thumbnails/large.jpg?file=", "");

      // console.log("For a stock image:");
      // console.log("Original Image Url:", image);
      // console.log("Modified Image Url:", imageUrl);
    } else if(image.includes("getpreview")){
      imageUrl = image.replace("_layouts/15/getpreview.ashx?path=%2F", "");

      // console.log("For a Uploaded image:");
      // console.log("Original Image Url:", image);
      // console.log("Modified Image Url:", imageUrl);
    }else{
      imageUrl = image;
    }
    return imageUrl;
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Url for See All"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
