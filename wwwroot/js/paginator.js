export default class Paginator{
    /**
     *  @param {Array} items ---> full list of items to paginate
     *  @param {Object} options
     *  @param {number} options.initialPageSize ---> how many items on page 0
     *  @param {number} options.chunkSize ---> how many items on pages >= 1
     *  @param {number} [options.startPage = 0] ---> initial page index
     */


    constructor(items, { initialPageSize, chunkSize, startPage = 0 })
    {
        
        this.items = Array.isArray(items) ? items : [];
        this.initialPageSize = Math.max(1 , initialPageSize);
        this.chunkSize = Math.max(1 , chunkSize);
        this.currentPage = Math.max(0 , startPage);
        this.totalItems = this.items.length;

    }


    // total pages (page0 + pages > 1)
    get TotalPages()
    {
        
        if(this.totalItems <= this.initialPageSize){
            return 1;
        }

        const rem = this.totalItems - this.initialPageSize;
        return 1 + Math.ceil(rem / this.chunkSize);

    }

    // To check if there's page after the current one
    get hasNext(){
        return this.currentPage < this.TotalPages - 1;
    }

    // To check if there's page before the current one
    get hasPrev(){
        return this.currentPage > 0;
    }


    // Cmpute start/end indices and slice out the current page
    get range(){
        
        let start , length;
        if(this.currentPage === 0){
            start = 0;
            length = this.initialPageSize;
        }
        else{
            start = this.initialPageSize + (this.currentPage - 1) * this.pageSize;
            length = this.pageSize;
        }

        const end = Math.min(start + length , this.totalItems);
        return {start , end , items : this.items.slice(start , end)};
    }

    
    // Moving to next page, return the new page items
    next(){
        if(this.hasNext){
            this.currentPage += 1;
        }
        return this.range.items;
    }


    // Move to the previous page, return the new page items
    prev(){
        if(this.hasPrev){
            this.currentPage -= 1;
        }
        return this.range.items;
    }


}