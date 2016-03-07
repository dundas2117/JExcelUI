// JavaScript source code


function ExcelUI(canvasId,eObj) {

    if (typeof (console) === "undefined") {

        console = {};

        console.log = function () { };

    }

    var c = document.getElementById(canvasId);
    var ctx = c.getContext("2d");

    var canvasWidth = c.width;
    var canvasHeight = c.height;

    //this object contains all informaiton needed to draw the UI

    var _colors = {
        hilightedBorder: "darkgreen",
        hilightedFill: "#E5E5E5",
        normalBorder: "#D3D3D3",
        normalFill:"white"
    };
    
    var _base = [1, 26, 676];
    var _eObj = eObj;

    if (_eObj == undefined) _eObj = new EObject();

    ctx.font = _eObj.font;

    var font = _eObj.font;

    var startXPos = 50;
    var startYPos = 20;
    
    //var width = 70;
    //var rowHeight = 20;

    var vScrollBarWidth = 20.0;
    var scrollButtonSize = 14.0;
    var hScrollBarHeight = 20.0;

    var firstVisibleRow = 1;
    var firstVisibleColumn = 1;
    var lastVisibleRow = 1;
    var lastVisibleColumn = 1;

  
    var controls = new Set();

    var _hilightedCells = new Set();


    var dragOk = false;
    var dragControl;

    var dataArea = {};
    dataArea.top = 1;
    dataArea.left = 1;
    dataArea.width = canvasWidth - vScrollBarWidth - 3;
    dataArea.height = canvasHeight - hScrollBarHeight - 1;
    dataArea.bottom = dataArea.top + dataArea.height - 2;

    var bottomBarArea = {};
    bottomBarArea.top = canvasHeight - hScrollBarHeight;
    bottomBarArea.height = hScrollBarHeight;
    bottomBarArea.left = 1;
    bottomBarArea.width = canvasWidth - vScrollBarWidth;

    var vScrollBarArea = {};
    vScrollBarArea.top = 1;
    vScrollBarArea.left = canvasWidth - vScrollBarWidth + 3;
    vScrollBarArea.width = vScrollBarWidth;
    vScrollBarArea.height = dataArea.height - 2;
   
    vScrollBarArea.scrollHeight = vScrollBarArea.height - 2 * scrollButtonSize;

    var scrollUpButton = new control("ScrollUpButton");
    scrollUpButton.setLeft(canvasWidth - vScrollBarWidth + 3);
    scrollUpButton.setTop(1);
    scrollUpButton.setWidth(scrollButtonSize);
    scrollUpButton.setHeight(scrollButtonSize);
    scrollUpButton.draw = function () {
        drawButton("U", scrollUpButton.getLeft(), scrollUpButton.getTop(), scrollUpButton.getLeft() + scrollUpButton.getWidth(), scrollUpButton.getTop() + scrollUpButton.getHeight(),
            scrollUpButton.getIsHighlighted());
    };
    scrollUpButton.click = scrollUp;


    var scrollDownButton = new control("ScrollDownButton");
    scrollDownButton.setLeft(canvasWidth - vScrollBarWidth + 3);
    scrollDownButton.setTop(vScrollBarArea.top + vScrollBarArea.height - scrollButtonSize);
    scrollDownButton.setWidth(scrollButtonSize);
    scrollDownButton.setHeight(scrollButtonSize);
    scrollDownButton.draw = function () {
        drawButton("D", scrollDownButton.getLeft(), scrollDownButton.getTop(), scrollDownButton.getLeft() + scrollDownButton.getWidth(), scrollDownButton.getTop() + scrollDownButton.getHeight(),
            scrollDownButton.getIsHighlighted());
    };
    scrollDownButton.click =scrollDown;

    var vScrollSlider = new control("VScrollSlider");
    vScrollSlider.setLeft(canvasWidth - vScrollBarWidth + 3);
    vScrollSlider.setTop(vScrollBarArea.top + vScrollBarArea.height - scrollButtonSize);
    vScrollSlider.setWidth(scrollButtonSize);
    //vScrollSlider.setHeight(scrollButtonSize);


    var vScrollRect = new control("vScrollRect");
    vScrollRect.setLeft(vScrollBarArea.left);
    vScrollRect.setTop(vScrollBarArea.top + scrollButtonSize);
    vScrollRect.setWidth(vScrollBarArea.width);
    vScrollRect.setHeight(vScrollBarArea.scrollHeight);
    vScrollRect.click = vScrollClick;
    vScrollRect.draw = function () { };

    

    var dragStartPos = { x: 0, y: 0 };

    init();

    function init()
    {
        
        controls.add(scrollUpButton);
        controls.add(scrollDownButton);
        controls.add(vScrollRect);


        c.addEventListener('mousemove', function (evt) {
            var mousePos = getMousePos(c, evt);
       
            if (!dragOk) {
                controls.forEach(function (item) {
                    if (item.isHit(mousePos.x, mousePos.y)) {

                        item.mouseover();
                        item.draw();
                    }
                    else {
                        item.mouseout();
                        item.draw();

                    }
                });
            }
            else {

                if (dragControl.getId() == "VScrollSlider") {
                    vScrollDelta(mousePos.y - dragStartPos.y);
                 
                  
                }

            }

           
            //console.log(message);
        }, false);

        c.addEventListener('mouseout', function (evt) {

            dragOk = false;
            controls.forEach(function (item) {
                if (item.getIsHighlighted()) {
                    item.mouseout();
                    item.draw();
                }
            });

         

        });

        c.addEventListener('click', function (evt) {
            var mousePos = getMousePos(c, evt);
            

            if (isPointInRect(mousePos, dataArea.left,dataArea.top, dataArea.width, dataArea.height)) {

                onDataAreaClick(mousePos.x, mousePos.y);

            }
            else {
                controls.forEach(function (item) {
                    if (item.isHit(mousePos.x, mousePos.y)) {
                        item.click(mousePos.x, mousePos.y);
                    }
                });

            }

        }, true);

        c.addEventListener('mousedown', function (evt) {
            var mousePos = getMousePos(c, evt);
            if (vScrollSlider.isHit(mousePos.x, mousePos.y)) {

                dragOk = true;
                dragStartPos.x = mousePos.x;
                dragStartPos.y = mousePos.y;
                dragControl = vScrollSlider;
                console.log("dragstart");
            }




        }, false);

        c.addEventListener('mouseup', function (evt) {
            var mousePos = getMousePos(c, evt);
          
            if (dragOk) {
                dragOk = false;
                dragControl = null;
            }




        }, false);

        c.addEventListener('mousewheel', function (evt) {
            

            vScrollDelta(evt.wheelDelta);
           



        }, false);

        firstVisibleRow = 1 + _eObj.frozenRows;
        draw(firstVisibleRow, firstVisibleColumn);
    }

    function onDataAreaClick(x, y) {

        var cell = getCellAtPoint(x, y);
        if (cell != null) {

            clickCell(cell);

        }

    }

    function clickCell(cell) {
        hilightCell(cell);

    }


    function getMousePos(canvas, evt) {
        var rect = canvas.getBoundingClientRect();
        return {
            x: evt.clientX - rect.left,
            y: evt.clientY - rect.top
        };
    }

    function draw(startRow, startColumn )
    {
        var counter = 1;
        var hLineLength =0;
        var vLineLength = 0;

      
        var yPos = startYPos;
        var xPos = startXPos;

        var rowHeight;
        var colWidth;

        ctx.clearRect(0, 0, canvasWidth, canvasHeight);

        
        //H grids, draw rows, no lines. lines will be drawn later
        //for frozen rows
        for (counter = 1; counter <= _eObj.frozenRows; counter++) {

            //get the height of the row to draw
            rowHeight = getRowHeight(counter);

            //draw row number
            drawTextCenter(counter, ctx, 0, yPos, startXPos, rowHeight, font, "black");

            //draw line seperating row number
            drawGradientLineH(ctx, 0, startXPos, yPos + rowHeight);

            vLineLength += rowHeight;
            yPos = yPos + rowHeight;

           
        }

        //for scrollable rows
        for (counter = 1; counter <= _eObj.numberOfRows - firstVisibleRow + 1 ; counter++) {

            //get the height of the row to draw
            rowHeight = getRowHeight(startRow + counter - 1);


            //draw row number
            drawTextCenter(startRow + counter - 1, ctx, 0, yPos, startXPos, rowHeight, font, "black");

            //draw line seperating row number
            drawGradientLineH(ctx, 0, startXPos, yPos + rowHeight);

            vLineLength += rowHeight;
            yPos = yPos + rowHeight;

            if (startYPos + vLineLength + 1 <= dataArea.height) {
                lastVisibleRow = startRow + counter - 1;  //this one
          
            }
            else
            {
                break;
            }
        }


        if (vLineLength > dataArea.height - startYPos) {
            vLineLength = dataArea.height - startYPos;
        }
        drawGradientLineV(ctx, startXPos, 0, startYPos);


     

        for (counter = 1; counter <= _eObj.numberOfColumn - firstVisibleColumn + 1 ; counter++) {

            colWidth = getColumnWidth(counter);

            drawTextCenter(String.fromCharCode(64 + startColumn + counter - 1), ctx, xPos, 0, colWidth, startYPos, font, "black");
           
            drawGradientLineV(ctx, xPos + colWidth, 0, startYPos);
            drawThinVerticalLine(ctx, xPos + colWidth, startYPos, startYPos + vLineLength, "#D3D3D3");

            hLineLength += colWidth;
            xPos = xPos + colWidth;

            //is it visible?
            if (startXPos + hLineLength + 1 <= dataArea.width) {
                lastVisibleColumn = startColumn + counter - 1;
            }
            else{
                break;
            }

        }



        if (hLineLength > dataArea.width- startXPos) {
            hLineLength = dataArea.width - startXPos;
        }

        drawGradientLineH(ctx, 0, startXPos, startYPos);


        //h grid line in frozen section
        yPos = startYPos;
        for (counter = 1; counter <= _eObj.frozenRows; counter++) {
            yPos = yPos + getRowHeight(counter);
            drawThinHorizontalLine(ctx, startXPos, startXPos + hLineLength, yPos, "#D3D3D3");
        }
        
        //h grid line in scrollable section
        for (counter = firstVisibleRow; counter <= lastVisibleRow ; counter++) {
            yPos = yPos + getRowHeight(counter);
            drawThinHorizontalLine(ctx, startXPos, startXPos + hLineLength, yPos, "#D3D3D3");
        }

        drawCombinedCell();
        drawData();
        drawHilightedCells();

        
        ctx.clearRect(dataArea.width, startYPos, dataArea.width + vScrollBarWidth, dataArea.height);
        
        drawBorders();
        
      
  
        drawVScrollBar();
        drawBottomBar();


      
        
    }

    //left/top/right/bottom border, and the frozen lines 
    function drawBorders() {
        //top border
        drawThinHorizontalLine(ctx, startXPos, dataArea.width, startYPos, "#808080");
        //bottom border
        drawThinHorizontalLine(ctx, startXPos, dataArea.width, dataArea.bottom, "#808080");

        //right border
        drawThinVerticalLine(ctx, dataArea.width, startYPos, dataArea.height, "#808080");

        //left border
        drawThinVerticalLine(ctx, startXPos, startYPos, dataArea.height, "#808080");

        if (_eObj.frozenRows != 0) {

            drawThinHorizontalLine(ctx, startXPos, dataArea.width, getCellPos({ row: _eObj.frozenRows, column: 1 }).y + getRowHeight(_eObj.frozenRows), "#808080");

        }

    }

    function drawHilightedCells() {

        _hilightedCells.forEach(function (item) {
            drawCell(item, true);
        });

    }
    // addr:  {row, column}
    function drawCell(cell, hilighted) {
        var pos = getCellPos(cell);

        var colWidth = getColumnWidth(cell.column);
        var rowHeight = getRowHeight(cell.row);

        var rectWidth;
        var rectHeight;


        var range = findCombinedRange(cell);

        if (range != null) {

            hilightCombinedCell(range, hilighted)

            return;
        }
            


        if (isColumnVisible(cell.column)) {
            hilightHeader(cell.column, hilighted);
        }

        if (isRowVisible(cell.row)) {
            hilightRowInd(cell.row, hilighted)
        }


        if (isCellVisible(cell)) {

          
            if (hilighted) {
                drawRectangle(ctx, pos.x, pos.y, colWidth, rowHeight, _colors.hilightedBorder);

            }
            else {


                drawRectangle(ctx, pos.x, pos.y, colWidth, rowHeight, _colors.normalBorder);
                //restore borders
                drawBorders();

            }
            



        }
       
        
        


    }

    function hilightHeader(col, hilighted) {
        var pos;
        pos = getCellPos({ row: firstVisibleRow, column: col });
        var colWidth = getColumnWidth(col);

        ctx.beginPath();
        //hilight column header
        ctx.rect(pos.x + 1, 0, colWidth - 1, startYPos);
        if (hilighted) {
            ctx.fillStyle = _colors.hilightedFill;
        }
        else {
            ctx.fillStyle = _colors.normalFill;
        }
        ctx.fill();
        //left border
        drawGradientLineV(ctx, pos.x, 0, startYPos, hilighted);
        //right border
        drawGradientLineV(ctx, pos.x + colWidth, 0, startYPos, hilighted);
        //add text 
        drawTextCenter(fromNumberToLetter(col), ctx, pos.x, 0, colWidth, startYPos, _eObj.font, "black");
    }

    function hilightRowInd(row, hilighted) {
        var pos;
        pos = getCellPos({ row: row, column: firstVisibleColumn });
        var rowHeight = getRowHeight(row);

        ctx.beginPath();
        //hilight column header
        ctx.rect(0, pos.y + 1, startXPos, rowHeight);
        if (hilighted) {
            ctx.fillStyle = _colors.hilightedFill;
        }
        else {
            ctx.fillStyle = _colors.normalFill;
        }
        ctx.fill();
        //top border
        drawGradientLineH(ctx, 0, startXPos, pos.y, hilighted);
        //bottom border
        drawGradientLineH(ctx, 0, startXPos, pos.y + rowHeight, hilighted);
        //add row number 
        drawTextCenter(row, ctx, 0, pos.y, startXPos, rowHeight, _eObj.font, "black");
    }
    //range: {startCell, endCell}
    function hilightCombinedCell(range, hilight) {
       
        var info = getCombinedRangeInfo(range)
        var borderColor = _colors.normalBorder;

        if (hilight) {
            borderColor = _colors.hilightedBorder;
        }

      

        if (info.topVisible) {
            drawThinHorizontalLine(ctx, info.left, info.left + info.width, info.top, borderColor);
        }

        if (info.leftVisible) {
            drawThinVerticalLine(ctx, info.left, info.top, info.top + info.height, borderColor);
        }

        if (info.bottomVisible) {
            drawThinHorizontalLine(ctx, info.left, info.left + info.width,info.top + info.height, borderColor);
        }

        if (info.rightVisible) {
            drawThinVerticalLine(ctx, info.left + info.width, info.top, info.top + info.height, borderColor);
        }

        //header (top )
        for (var c = info.startCell.column; c <= info.endCell.column; c++) {

            if (isColumnVisible(c)){
                hilightHeader(c, hilight);
            }

        }

        //row indicator
        for (var r = info.startCell.row; r <= info.endCell.row; r++) {
            if (isRowVisible(r)) {
                hilightRowInd(r, hilight);
            }
        }



    }

    //return {startCell, endCell} if the cell {row, column} is 
    //in a combined area.
    //otherwise return null
    function findCombinedRange(cell) {
        var pair;
        var startCell;
        var endCell;
        var found = false;

        _eObj.combinedCells.forEach(function (item) {

            pair = item.split(":");
            startCell = getCellAddress(pair[0]);
            endCell = getCellAddress(pair[1]);

            if (cell.row >= startCell.row && cell.row <= endCell.row && cell.column >= startCell.column && cell.column <= endCell.column) {
                found = true;
                return false;
            }
        });

        if (found)
        {
            return {startCell:startCell, endCell:endCell};
        }
        else {

            return null;
        }

    }


    function hilightCell(cell) {

        _hilightedCells.forEach(function (item) {
            drawCell(item, false);
        });
        _hilightedCells.clear();
        _hilightedCells.add(cell);
        drawHilightedCells();
    }


   
    

    //return AA for 27
    function fromNumberToLetter(n){
        var r ="";
        var q=0;

        for ( var i = _base.length-1; i>=0; i--){
            q = Math.floor(n/_base[i]);
            if (q >=1){
                r = r+ String.fromCharCode(64 + q);
                n = n- q*_base[i];
            }
        }

        return r;

      
    }

    function drawData() {
        var counter = 1;
        var pos;
        _eObj.data.forEach(function (item, key, mapObj) {
            var addr = getCellAddress(key);
            if (isCellVisible(addr)) {

                pos = getCellPos(addr);
                    
                drawTextCenter(item, ctx, pos.x, pos.y, getColumnWidth(addr.column), getRowHeight(addr.row), font, "black");

            }
        });


    }


    
    //draw combined cells, simple clear the inner grids for now
    function drawCombinedCell() {

        var pair;
        var startCell;
        var endCell;

        _eObj.combinedCells.forEach(function (item) {
            
            pair = item.split(":");
            startCell = getCellAddress(pair[0]);
            endCell = getCellAddress(pair[1]);
            var w = 0;
            var startCellPos = getCellPos(startCell);
            var yPos ;
            var height;

            //combined cell not visible, ignore it.
            if (startCell.row > _eObj.frozenRows && (startCell.row > lastVisibleRow || endCell.row < firstVisibleRow)) return;
            
            //clear the inner grids
            for (var c = startCell.column; c <= endCell.column; c++) {
                w = w + getColumnWidth(c);

            }

            if (w + startCellPos.x > dataArea.width) {

                w = dataArea.width - startCellPos.x;
            }


            for (var r = startCell.row; r <= endCell.row; r++) {
                
                if (r > _eObj.frozenRows && r < firstVisibleRow) continue;
                yPos = getCellPos({row:r,column:1}).y;
                height = getRowHeight(r);
                if ( r == startCell.row) yPos = yPos + 1;
               
                ctx.clearRect(startCellPos.x + 1, yPos ,w - 1 , height   );
            }





        });

    }

    //parameter range {startCell, endCell}
    //return combinedRangeInfo
    function getCombinedRangeInfo(range) {
        var info = new combinedRnageInfo();
        var startCell = range.startCell;
        var endCell = range.endCell;
        var w = 0;
        var startCellPos = getCellPos(startCell);
        var pos;
        var cell;

        var columnsCounted = [];
        var rowsCounted = [];
        
        

        info.startCell = startCell;
        info.endCell = endCell;

        

        info.topVisible = !(startCell.row > _eObj.frozenRows && startCell.row < firstVisibleRow )

        info.bottomVisible = !(endCell.row > lastVisibleRow);

        info.leftVisible = !(startCell.column > _eObj.frozenColumns && startCell.column < firstVisibleColumn);
       
        info.rightVisible = !(endCell.column > lastVisibleColumn);

        info.visible = (info.leftVisible || info.rightVisible || info.topVisible || info.bottomVisible);


        for (var r = startCell.row; r <= endCell.row; r++) {

          
            for (var c = startCell.column; c <= endCell.column; c++) {
                cell = {row:r,column:c};
                if (isCellVisible(cell)) {

                    if (info.startVisibleCell.row ==0){
                        info.startVisibleCell.row = r;
                        info.startVisibleCell.column =c;
                    }

                    info.endVisibleCell.row = r;
                    info.endVisibleCell.column = c;

                    pos = getCellPos(cell);
                    if (info.left == 0) {
                        info.left = pos.x;
                        info.top = pos.y;
                    }
                    if (columnsCounted.indexOf(c) < 0) {
                        info.width = info.width + getColumnWidth(c);
                        columnsCounted.push(c);
                    }
                    if (rowsCounted.indexOf(r) < 0) {
                        info.height = info.height + getRowHeight(r);
                        rowsCounted.push(r);
                    }
                }
            }
        }

        return info;

    }

   
    //Get list of cells ([row, column]) in the range (A1:B20);
    function getCellsInRange(range) {
        var pair;
        var startCell;
        var endCell;
        var list = [];

        pair = item.split(":");
        startCell = getCellAddress(pair[0]);
        endCell = getCellAddress(pair[1]);

        for (var r = startCell.row ; r <= endCell.row; r++) {
            for (var c = startCell.column; c <= endCell.column; c++) {
                list.push({ row: r, column: c });

            }

        }

        return list;

    }

    //return cell {row,column} at the point
    function getCellAtPoint(x, y) {
        
        var cell;
        
        cell = getCellAtPointInARange(x,y,{ row: 1, column: firstVisibleColumn }, { row: _eObj.frozenRows, column: lastVisibleColumn });
        if (cell == null) {

            cell = getCellAtPointInARange(x,y,{ row: firstVisibleRow, column: firstVisibleColumn }, { row: lastVisibleRow, column: lastVisibleColumn });
        }


        return cell;


    }

    function getCellAtPointInARange(x, y, firstCell, lastCell) {
        var found = false;
        var pos;
        var r = 0;
        var c = 0;

        for (r = firstCell.row; r <= lastCell.row ; r++) {

            for (c = firstCell.column; c <= lastCell.column; c++) {

                pos = getCellPos({ row: r, column: c });
                if (isPointInRect({ x: x, y: y }, pos.x, pos.y, getColumnWidth(c), getRowHeight(r))) {

                    found = true;
                    break;
                }

            }

            if (found) break;

        }

        if (found) {
            return { row: r, column: c };
        }
        else
        {
            return null;
        }

    }

    //return if a point in in a rect or not.
    function isPointInRect(pos, left, top, width, height) {

        return (pos.x >= left && pos.x <= left + width && pos.y >= top && pos.y <= top + height);
    }

    //addr:  (row, column)
    //return pos {x:xpos, y:ypos}
    function getCellPos(addr) {

        var yPos = startYPos;
        var xPos = startXPos;


        if (addr.row <= _eObj.frozenRows) {
            for (var i = 1; i < addr.row ; i++) {
                yPos = yPos + getRowHeight(i);
            }
        }
        else {
            yPos = yPos + getFrozenHeigth();
            for (var i = firstVisibleRow; i < addr.row; i++) {
                yPos = yPos + getRowHeight(i);
            }
        }

        if (addr.column <= _eObj.frozenColumns) {
            for (var i = 1; i < addr.column ; i++) {
                xPos = xPos + getColumnWidth(i);
            }
        }
        else {
            xPos = xPos + getFrozenWidth();
            for (var i = firstVisibleColumn; i < addr.column; i++) {
                xPos = xPos + getColumnWidth(i);
            }
        }

        return { x: xPos, y: yPos };

    }

    function getFrozenHeigth() {
        var height = 0;
        for (var i = 1; i <= _eObj.frozenRows ; i++) {
            height = height + getRowHeight(i);
        }

        return height;
    }

    function getFrozenWidth() {
        var width = 0;
        for (var i = 1; i < _eObj.frozenColumns; i++) {
            width = width + getColumnWidth(i);

        }

        return width;
    }


    function getRowHeight(row) {
        var val ;
        if (_eObj.rowStyles.has(row)) {
            val = _eObj.rowStyles.get(row).rowHeight;
        }
        if (val == undefined) val = 20;
        return val;
    }

    function getColumnWidth(col) {
        var val = 70;
        if (_eObj.columnStyles.has(col)) {
            val = _eObj.columnStyles.get(col).columnWidth;
        }

        if (val == undefined) val = 70;

        return val;
    }


    // return if a cell ({row, column} is visible or not.
    function isCellVisible(addr){
        
        var rowVisible = false;
        var colVisible = false;

        rowVisible = isRowVisible(addr.row);

        if (rowVisible){
            colVisible = isColumnVisible(addr.column);
        }

        return rowVisible && colVisible;

        
        
    }

    //find out if a column is visible or not. 
    //return true/false
    //isColumnVisible(3)
    function isColumnVisible(col) {
        var visible = false;
        if (col <= _eObj.frozenColumns) {
            visible  = true;
        }
        else if (col >= firstVisibleColumn &&col <= lastVisibleColumn + 1) {
            visible = true;
        }

        return visible;
    }

    //find out if a row is visible or not. 
    //return true/false
    //isRowVisible(3)
    function isRowVisible(row) {

        var visible = false;

        if (row <= _eObj.frozenRows) {
            visible = true;
        }
        else{

            if (row >= firstVisibleRow && row <= lastVisibleRow + 1) {
                visible = true;
            }
        }

        return visible;
    }

    //return {row, column} from address such as A128
    function getCellAddress(addressName) {
        var rowPart;
        var row = 0;
        var colPart;
        var col = 0;

      

        for (var i = 0; i < addressName.length; i++) {
            if (!isNaN(parseInt(addressName.substr(i, 1)))) {
                rowPart = addressName.substr(i, addressName.length - i);
                row = parseInt(rowPart);
                break;
            }

        }
        colPart = addressName.substr(0, i );
        for (var i = colPart.length - 1; i >= 0; i--) {
            col = col + (colPart.charCodeAt(i) - 64) * _base[i - colPart.length + 1];

        }

        return { row: row, column: col };

    }

    

    function drawThinHorizontalLine(c, x1, x2, y, style) {
        c.lineWidth = 1;
        var adaptedY = Math.floor(y) + 0.5;
        c.beginPath();
        c.moveTo(x1, adaptedY);
        c.lineTo(x2, adaptedY);
        c.strokeStyle = style;
        c.stroke();
    }
    function drawThinVerticalLine(c, x, y1, y2, style) {
        c.lineWidth = 1;
        var adaptedX = Math.floor(x) + 0.5;
        c.beginPath();
        c.moveTo(adaptedX, y1);
        c.lineTo(adaptedX, y2);
        c.strokeStyle = style;
        c.stroke();
    }

    function drawGradientLineH(c, x1, x2, y,hilighted) {
        var grad = c.createLinearGradient(x1, y, x2, y);
        if (hilighted) {
            grad.addColorStop(0, "#D3D3D3");
            grad.addColorStop(1, "#808080");
        }
        else {
            grad.addColorStop(0, "white");
            grad.addColorStop(1, "#D3D3D3");
        }
        drawThinHorizontalLine(c, x1, x2, y, grad);
    }

    function drawGradientLineV(c, x, y1, y2, hilighted) {
        var grad = c.createLinearGradient(x, y1, x, y2);
        if (hilighted) {
            grad.addColorStop(0, "#D3D3D3");
            grad.addColorStop(1, "#808080");
        }
        else {
            grad.addColorStop(0, "white");
            grad.addColorStop(1, "#D3D3D3");
        }
        drawThinVerticalLine(c, x, y1, y2, grad);
    }

  

    function drawTextCenter(text, c, x1, y1, width, height, font, color) {
        var x = x1 + width / 2;
        var y = y1 + height / 2;
        c.font = font;
        c.textAlign = "center";
        c.textBaseline = "middle";
        c.fillStyle = color;
        c.fillText(text, x, y);

    }

    function scrollDown(){
    
        if (lastVisibleRow == firstVisibleRow) return;
        firstVisibleRow = firstVisibleRow + 1;
        draw(firstVisibleRow, firstVisibleColumn);
    }

    function scrollUp() {

        if (firstVisibleRow == 1 + _eObj.frozenRows) return;
        firstVisibleRow = firstVisibleRow - 1;
        draw(firstVisibleRow, firstVisibleColumn);
    }

    function scrollRight(){
        if (lastVisibleColumn == _eObj.numberOfColumn) return;
        firstVisibleColumn++;
        draw(firstVisibleRow, firstVisibleColumn);
    }

    function scrollLeft() {
        if (firstVisibleColumn == 1) return;
        firstVisibleColumn--;
        draw(firstVisibleRow, firstVisibleColumn);
    }

    function vScrollDelta(delta) {

       
        
        var rows = Math.floor(delta * (_eObj.numberOfRows - _eObj.frozenRows )/ vScrollBarArea.scrollHeight);
        firstVisibleRow = rows + firstVisibleRow;
        if (firstVisibleRow < 1 + _eObj.frozenRows) {
            firstVisibleRow = 1 + _eObj.frozenRows;
        }
        if (firstVisibleRow > _eObj.numberOfRows) {
            firstVisibleRow = _eObj.numberOfRows;
        }

      
        
        draw(firstVisibleRow, firstVisibleColumn);

        



    }

    function vScrollClick(x,y) {

        var delta = 0;
        if (y > vScrollSlider.getTop() + vScrollSlider.getHeight()) {
            delta = y - vScrollSlider.getTop() - vScrollSlider.getHeight();
        }

        if (y < vScrollSlider.getTop()) {
            delta = y - vScrollSlider.getTop();
        }
        vScrollDelta(delta);

    }

    function drawButton(flag, left,top,right,bottom, hilighted) {

        ctx.clearRect(left, top, right, bottom);
        var color = "#808080";

        if (hilighted) color = "black";

        drawRectangle(ctx, left, top, right - left, bottom - top,color);


        if (flag == "U") {
            //draw up arrow
            drawArrow(ctx, left + scrollButtonSize / 2.0 + 0.5, top +4,
                            left + 3, bottom -5,
                            right -2 , bottom - 5,
                            "gray"
                            );

        }
        else if (flag=="D") {
            drawArrow(ctx, left + scrollButtonSize / 2.0 + 0.5, bottom - 4,
                           left + 3, top + 5 ,
                           right - 2, top + 5 ,
                           "gray"
                           );
        }
        else if (flag == "L") {

            drawArrow(ctx, left + 3, top + scrollButtonSize/2,
                                      right -5 , top + 3,
                                      right - 5, bottom - 2,
                                      "gray"
                                      );
        }
        else if (flag == "R") {

            drawArrow(ctx, left + 5, top + 3,
                                      right - 3, top + scrollButtonSize / 2,
                                      left+5, bottom - 2,
                                      "gray"
                                      );
        }

    }

    function drawArrow(ctx,x1,y1,x2,y2,x3,y3,color) {
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.lineTo(x3, y3);
        ctx.closePath();
        ctx.fill();

    }

    function drawRectangle(ctx, x1, y1, width, height, color) {

        drawThinHorizontalLine(ctx, x1, x1 + width, y1,color);
        drawThinHorizontalLine(ctx, x1, x1+width, y1 + height,color);

        drawThinVerticalLine(ctx, x1, y1, y1+height,color);
        drawThinVerticalLine(ctx, x1 + width, y1, y1 + height+1, color);

    }


    function drawVScrollBar() {
        //now draw vscrollbar



        //drawButton("U", left, top, right, bottom);
        scrollUpButton.draw();

        scrollDownButton.draw();

        //draw slider
        drawVScrollArea();
    }

    function drawVScrollArea() {
        vScrollSlider.setHeight (Math.floor((lastVisibleRow - firstVisibleRow + 1) / (_eObj.numberOfRows - _eObj.frozenRows) * vScrollBarArea.scrollHeight));

        ctx.clearRect(vScrollBarArea.left, vScrollBarArea.top + scrollButtonSize, vScrollBarArea.left + vScrollBarArea.width, scrollDownButton.top - 1);

        vScrollSlider.setTop( Math.floor(vScrollBarArea.top + scrollButtonSize  + 1 + (firstVisibleRow - 1 - _eObj.frozenRows) / (_eObj.numberOfRows-_eObj.frozenRows) * vScrollBarArea.scrollHeight));
        drawRectangle(ctx, vScrollBarArea.left, vScrollSlider.getTop(), scrollButtonSize, vScrollSlider.getHeight());

    }

    
    function drawBottomBar() {
        //clear bottom 
        ctx.clearRect(bottomBarArea.left, bottomBarArea.top, bottomBarArea.left + bottomBarArea.width, bottomBarArea.top + bottomBarArea.height);

       
      
        
        drawHScrollBar();
    }

    function drawHScrollBar() {
        var left = canvasWidth / 2;
        var right = canvasWidth / 2 + scrollButtonSize;
        var top = canvasHeight - scrollButtonSize - 1;
        var bottom = canvasHeight - 1;

        drawButton("L", left, top, right, bottom);

        left = dataArea.left + dataArea.width - scrollButtonSize;
        right = left + scrollButtonSize;

        drawButton("R", left, top, right, bottom);


    }

 

    return {
    
        scrollUp: scrollUp,
        scrollDown: scrollDown,
        scrollLeft: scrollLeft,
        scrollRight: scrollRight,
        hilightCell: hilightCell


      
        

    };




}

function control(id) {
    var left;
    var top;
    var width;
    var height;

    var Id;
    var isHighlighted = false;

    var isDragable = false;

    Id = id;



    return {

        getLeft: function () { return left; },
        getWidth: function () { return width; },
        getTop: function () { return top; },
        getHeight: function () { return height },
        getIsHighlighted: function () { return isHighlighted; },
        getIsDragable : function() { return isDragable},

        setLeft: function (v) { left = v; },
        setTop: function (v) { top = v; },
        setWidth: function (v) { width = v; },
        setHeight: function (v) { height = v; },
        setIsHighlighted: function (v) { isHighlighted = v; },
        setIsDragable: function (v) { isDragable = v; },

        getId: function () { return Id; },

        isHit: function (x, y) {
            return (x >= left && x <= left + width && y >= top && y <= top + height);

        },



        mouseover: function () {
            isHighlighted = true;
           

        },
        mouseout: function () {
            isHighlighted = false;
        }

    }

}
function EObject() {

    var columnStyles = new Map();
    var rowStyles = new Map();

    var data = new Map();
    var combinedCells = new Set();


        return {
            numberOfRows: 100,
            numberOfColumn: 26,
            font: "11pt Calibri",
            frozenRows: 0,
            frozenColumns: 0,

            columnStyles: columnStyles,
            rowStyles: rowStyles,

            combinedCells:combinedCells, //not supported yet

            data:data

        }
}

function combinedRnageInfo() {
    return {
        startCell:{row:0,column:0},
        endCell:{row:0,column:0},
        left: 0,
        top: 0,
        width: 0,
        height:0,
        visible:true,
        leftVisible:true,
        rightVisible:true,
        topVisible:true,
        bottomVisible:true,
        startVisibleCell: {row:0,column:0},
        endVisibleCell:{row:0,column:0}

    }
}

//rowStyle = {rowHeight:100,font:''}
//columnStyles={columnWidth:100,font:''}