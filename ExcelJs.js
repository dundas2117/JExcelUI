// JavaScript source code

(function (ExcelUINS, $, undefined) {
    function ExcelUI(canvasId,eObj) {

        if (typeof (console) === "undefined") {

            console = {};

            console.log = function () { };

        }

        if (typeof StopIteration == "undefined") {
            StopIteration = new Error("StopIteration");
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
            normalFill:"white",
            areaBorder: "#808080",
            scrollArea: "WhiteSmoke"

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
        var _lastMouseOver = null;

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
        bottomBarArea.scrollWidth = (canvasWidth / 2 - 2 * scrollButtonSize - vScrollBarWidth);

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
        scrollDownButton.click = scrollDown;


        var leftScrollButton = new control("LeftScrollButton");
        leftScrollButton.setLeft(canvasWidth / 2);
        leftScrollButton.setTop(canvasHeight - scrollButtonSize - 1);
        leftScrollButton.setWidth(scrollButtonSize);
        leftScrollButton.setHeight(scrollButtonSize);
        leftScrollButton.draw = function () {
            drawButton("L", leftScrollButton.getLeft(), leftScrollButton.getTop(), leftScrollButton.getLeft() + leftScrollButton.getWidth(), leftScrollButton.getTop() + leftScrollButton.getHeight(),
                leftScrollButton.getIsHighlighted());
        };
        leftScrollButton.click = scrollLeft;

        var rightScrollButton = new control("RightScrollButton");
        rightScrollButton.setLeft(vScrollBarArea.left - vScrollBarWidth );
        rightScrollButton.setTop(canvasHeight - scrollButtonSize - 1);
        rightScrollButton.setWidth(scrollButtonSize);
        rightScrollButton.setHeight(scrollButtonSize);
        rightScrollButton.draw = function () {
            drawButton("R", rightScrollButton.getLeft(), rightScrollButton.getTop(), rightScrollButton.getLeft() + rightScrollButton.getWidth(), rightScrollButton.getTop() + rightScrollButton.getHeight(),
                rightScrollButton.getIsHighlighted());
        };
        rightScrollButton.click = scrollRight;

        var hScrollSlider = new control("HScrollSlider");
        hScrollSlider.setLeft(canvasWidth / 2 + scrollButtonSize+1);
        hScrollSlider.setTop(leftScrollButton.getTop());
        hScrollSlider.setWidth(scrollButtonSize);
        //vScrollSlider.setHeight(scrollButtonSize);


        var hScrollRect = new control("HScrollRect");
        hScrollRect.setLeft(canvasWidth / 2 + scrollButtonSize);
        hScrollRect.setTop(canvasHeight - hScrollBarHeight);
        hScrollRect.setWidth(bottomBarArea.scrollWidth );
        hScrollRect.setHeight(hScrollBarHeight);
        hScrollRect.click = hScrollClick;
        hScrollRect.draw = function () { };



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
            controls.add(leftScrollButton);
            controls.add(rightScrollButton);
            controls.add(hScrollRect);


            c.addEventListener('mousemove', function (evt) {
                var mousePos = getMousePos(c, evt);
       
                if (!dragOk) {
                    try{
                        controls.forEach(function (item) {
                            if (item.isHit(mousePos.x, mousePos.y)) {
                                if (_lastMouseOver != null) {

                                    if (_lastMouseOver != item) {
                                        _lastMouseOver.mouseout();
                                        _lastMouseOver.draw();
                                    }
                                    else {
                                        item.mouseover();
                                        item.draw();
                                        _lastMouseOver = item;
                                    }
                                }
                                else {
                                    item.mouseover();
                                    item.draw();
                                    _lastMouseOver = item;

                                }

                                throw StopIteration;
                            }
                       
                  
                        });

                        if (_lastMouseOver != null) {
                            _lastMouseOver.mouseout();
                            _lastMouseOver.draw();
                            _lastMouseOver = null;
                        }
                    
                    }
                    catch (error) {
                        if (error != StopIteration) throw error;
                    }
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
                try{
                    controls.forEach(function (item) {
                        if (item = _lastMouseOver) {
                            item.mouseout();
                            item.draw();
                            _lastMouseOver = null;
                            throw StopIteration;
                        }
                    });
             
                }
                catch (error) {
                    if (error != StopIteration) throw error;
                }

         

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
                            return false;
                        }
                    });

                }

            }, true);

            c.addEventListener('dblclick', function (evt) {
                var mousePos = getMousePos(c, evt);


                if (isPointInRect(mousePos, dataArea.left, dataArea.top, dataArea.width, dataArea.height)) {

                    onDataAreaDoubleClick(mousePos.x, mousePos.y);

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
            
                var delta = evt.wheelDelta;
                var row = 0;
                if (delta > 0) {
                    row = -1;
                }
                else
                {
                    row = 1;
                }
                vScrollRows(row);
           



            }, false);

            firstVisibleRow = 1 + _eObj.frozenRows;
            firstVisibleColumn = 1 +_eObj.frozenColumns;
            draw(firstVisibleRow, firstVisibleColumn);
            //draw(firstVisibleRow, 27);
        }

        function onDataAreaDoubleClick(x, y) {

            var cell = getCellAtPoint(x, y);
            if (cell != null) {

                doubleClickCell(cell);

            }

        }

        //cell: {row, column}
        function doubleClickCell(cell) {
            var top ;
            var left;
            var width;
            var height;

            var pos = getCellPos(cell);
            top =  pos.y.toString() + 'px' ;
            left = pos.x.toString() + 'px';

            /*
            var prevCtrl = $("#cellinput");
            if (prevCtrl != null) {
                prevCtrl.remove();
            }
            */

            width = (getColumnWidth(cell.column) - 1);
            height = (getRowHeight(cell.row)-1);

            var thisCellStyle = getCellStyle(cell);

            var data = getCellData(cell);

            var $ctrl = getInputForCell(thisCellStyle,data,cell,left,top, width,height);

          
            $("#container").append($ctrl);
            $ctrl.focus();
        }

        //cellStyle: cellStyle
        function getInputForCell(cellStyle, data, cell,left, top, width,height) {
            var ctrl;
            switch (cellStyle.input) {

                case inputEnum.singleLine:
                    ctrl = $("<input id='cellinput'/>");
                    if (cellStyle.hAlign == hAlignEnum.right) {
                        ctrl.css("text-align", "right");

                    }
                    else if (cellStyle.hAlign == hAlignEnum.left) {
                        ctrl.css("text-align", "left");
                    }
                    if (data != null) {
                        ctrl.val(data);
                    }

                    ctrl.blur(function () {
                        setCellData(cell, $(this).val())
                        $(this).remove();
                        drawCellData($(this).val(), cell);
                    });
                    ctrl.css({ position: 'absolute', top: top, left: left, width: width.toString() + 'px', height: height.toString() + 'px', margin: 0, padding: 0 }).css("z-index", 100);
                    break;
                case inputEnum.multipleLine:

                    ctrl = $("<textarea id='cellinput'></textarea");
                    if (data != null) {
                        ctrl.text(data);
                    }
                    ctrl.blur(function () {
                        setCellData(cell, $(this).text());
                        $(this).remove();
                        drawCellData($(this).text(), cell);
                    });
                    ctrl.css({ position: 'absolute', top: top, left: left, width: width.toString() + 'px', height: height.toString() + 'px', margin: 0, padding: 0 }).css("z-index", 100);
                    break;

                case inputEnum.dropdownList:
                    ctrl = $("<select id='cellinput'><option>F5</option><option>F5-2</option></select>");
                    if (data != null) {
                        ctrl.val(data);
                    }
                    ctrl.blur(function () {
                        setCellData(cell, $(this).val());
                        $(this).remove();
                        drawCellData($(this).val(), cell);
                    });
                    ctrl.css({ position: 'absolute', top: top, left: left, width: (width + 2).toString() + 'px', height: (height + 2).toString() + 'px', margin: 0, padding: 0 }).css("z-index", 100);
                    break;
            }
  
            return ctrl;
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


            //H grids, draw rows, no lines. lines will be drawn later
            //for frozen rows
            for (counter = 1; counter <= _eObj.frozenColumns; counter++) {

                colWidth = getColumnWidth(counter);


                drawTextCenter(fromNumberToLetter( counter ), ctx, xPos, 0, colWidth, startYPos, font, "black");

                drawGradientLineV(ctx, xPos + colWidth, 0, startYPos);
                drawThinVerticalLine(ctx, xPos + colWidth, startYPos, startYPos + vLineLength, "#D3D3D3");

                hLineLength += colWidth;
                xPos = xPos + colWidth;


            }

            for (counter = 1; counter <= _eObj.numberOfColumn - firstVisibleColumn + 1 ; counter++) {

                colWidth = getColumnWidth(startColumn + counter - 1);

                drawTextCenter(fromNumberToLetter( startColumn + counter - 1), ctx, xPos, 0, colWidth, startYPos, font, "black");
           
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
            ctx.clearRect(dataArea.left, dataArea.bottom, dataArea.width , hScrollBarHeight);
        
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

            if (_eObj.frozenColumns != 0) {

                drawThinVerticalLine(ctx, getCellPos({row:1,column:_eObj.frozenColumns}).x + getColumnWidth(_eObj.frozenColumns), startYPos, dataArea.height, "#808080");

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
       
            var info = getCombinedRangeInfo(range);
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
        function fromNumberToLetter(n) {
            var r = "";
            var q = 0;
            var a;

            while (n > 0) {
                a = n % 26;
                if (a == 0) {
                    r = "Z" + r;
                    a = 26;
                }
                else {
                    r = String.fromCharCode(64 + a) + r;
                }

                n = Math.floor((n - a) / 26);

            }


            return r;

        }



        //return 27 for AA
        function fromLetterToNumber(s) {

            var col = 0;
            var j = 0;
            for (var i = s.length - 1; i >= 0; i--) {
                col = col + (s.charCodeAt(i) - 64) * _base[j];
                j++;




            }




            return col;




        }

        function drawData() {
            var counter = 1;
            var pos;
            var style;

            _eObj.data.forEach(function (item, key, mapObj) {
                var cell = getCellAddress(key);
                if (isCellVisible(cell)) {

               
                    drawCellData(item, cell)
               
                }
            });


        }

        function drawCellData(text,  cell) {

            var pos = getCellPos(cell);
            var style = getCellStyle(cell);
            var width = getColumnWidth(cell.column);
            var height = getRowHeight(cell.row);

            ctx.clearRect(pos.x + 1, pos.y + 1, width - 1, height - 1);

            var s = fittingString(ctx, text, width);

            if (style.hAlign == hAlignEnum.centre) {
                drawTextCenter(s, ctx, pos.x, pos.y, width, height, style.font, style.color);
            }
            else if (style.hAlign == hAlignEnum.right) {

                drawTextRight(s, ctx, pos.x, pos.y, width, height, style.font, style.color);

            }
            else {
                drawTextLeft(s, ctx, pos.x + 2, pos.y, width, height, style.font, style.color);
            }
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
           

                var info = getCombinedRangeInfo({ startCell: startCell, endCell: endCell });
                var startVisiblePos = getCellPos(info.startVisibleCell);
                var endVisiblePos = getCellPos(info.endVisibleCell);

                if (!info.visible) return;
                ctx.clearRect(startVisiblePos.x + 1, startVisiblePos.y + 1, endVisiblePos.x - startVisiblePos.x + getColumnWidth(info.endVisibleCell.column) - 1, endVisiblePos.y - startVisiblePos.y + getRowHeight(info.endVisibleCell.row) - 1);

        



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
        
            cell = getCellAtPointInARange(x, y, { row: 1, column: 1 }, { row: _eObj.frozenRows, column: lastVisibleColumn });

            if (cell == null) {
                cell = getCellAtPointInARange(x, y, { row: 1, column: 1 }, { row: lastVisibleRow, column: _eObj.frozenColumns });
            }
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

        //cell: {row, column}
        //return data for that cell
        function getCellData(cell) {
       
            var addrName = getCellAddressName(cell);
            var data;

            if (_eObj.data.has(addrName)) {

                data = _eObj.data.get(addrName);
            }

            return data;

        }

        //cell: {row, column}
        //return cellStyle for that cell
        function getCellStyle(cell) {
            var addrName = getCellAddressName(cell);
        
            return getCellStyleByAddrName(addrName);
        }

        //addrName: A128
        //return cellStyle
        function getCellStyleByAddrName(addrName) {
   
            var style = new cellStyle();

            if (_eObj.cellStyles.has(addrName)) {
                style = _eObj.cellStyles.get(addrName);
            }

            return style;
        }

        function setCellData(cell, data) {

            var addrName = getCellAddressName(cell);
            var data;

            _eObj.data.set(addrName, data);
      

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
            for (var i = 1; i <= _eObj.frozenColumns; i++) {
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
            col = fromLetterToNumber(colPart);
            return { row: row, column: col };

        }

        //cell : {row, column}
        //return address such as A128
        function getCellAddressName(cell) {
            var colName = fromNumberToLetter(cell.column);
            var addrName = colName + cell.row.toString();

            return addrName;
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

        function drawTextRight(text, c, x1, y1, width, height, font, color) {
            var x = x1 + width;
            var y = y1 + height / 2;
            c.font = font;
            c.textAlign = "right";
            c.textBaseline = "middle";
            c.fillStyle = color;
            c.fillText(text, x, y);

        }

        function drawTextLeft(text, c, x1, y1, width, height, font, color) {
            var x = x1 ;
            var y = y1 + height / 2;
            c.font = font;
            c.textAlign = "left";
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
            if (firstVisibleColumn == 1 + _eObj.frozenColumns) return;
            firstVisibleColumn--;
            draw(firstVisibleRow, firstVisibleColumn);
        }

        //scroll number of rows
        function vScrollRows(rows) {
            firstVisibleRow = rows + firstVisibleRow;
            if (firstVisibleRow < 1 + _eObj.frozenRows) {
                firstVisibleRow = 1 + _eObj.frozenRows;
            }
            if (firstVisibleRow > _eObj.numberOfRows) {
                firstVisibleRow = _eObj.numberOfRows;
            }



            draw(firstVisibleRow, firstVisibleColumn);
        }

        function vScrollDelta(delta) {

       
        
            var rows = Math.floor(delta * (_eObj.numberOfRows - _eObj.frozenRows )/ vScrollBarArea.scrollHeight);
          
            vScrollRows(rows);



        }

        function hScrollDelta(delta) {



            var columns = Math.floor(delta * (_eObj.numberOfColumn - _eObj.frozenColumns) / bottomBarArea.scrollWidth);
            firstVisibleColumn = columns + firstVisibleColumn;
            if (firstVisibleColumn < 1 + _eObj.frozenColumns) {
                firstVisibleColumn = 1 + _eObj.frozenColumns;
            }
            if (firstVisibleColumn > _eObj.numberOfColumn) {
                firstVisibleColumn = _eObj.numberOfColumn;
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

        function hScrollClick(x, y) {

            var delta = 0;
            if (x > hScrollSlider.getLeft() + hScrollSlider.getWidth()) {
                delta = x - hScrollSlider.getLeft() - hScrollSlider.getWidth();
            }

            if (x < hScrollSlider.getLeft()) {
                delta = x - hScrollSlider.getLeft();
            }
            hScrollDelta(delta);

        }


        function drawButton(flag, left,top,right,bottom, hilighted) {

            ctx.clearRect(left, top, right-left, bottom-top);
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

            ctx.clearRect(vScrollBarArea.left, vScrollBarArea.top + scrollButtonSize + 1, vScrollBarArea.width, vScrollBarArea.scrollHeight-1);
            ctx.fillStyle = _colors.scrollArea;
            ctx.fillRect(vScrollBarArea.left, vScrollBarArea.top + scrollButtonSize + 1, scrollButtonSize + 1, vScrollBarArea.scrollHeight-1)

            vScrollSlider.setTop( Math.floor(vScrollBarArea.top + scrollButtonSize  + 1 + (firstVisibleRow - 1 - _eObj.frozenRows) / (_eObj.numberOfRows-_eObj.frozenRows) * vScrollBarArea.scrollHeight));
            drawRectangle(ctx, vScrollBarArea.left, vScrollSlider.getTop(), scrollButtonSize, vScrollSlider.getHeight(),_colors.areaBorder);

        }

    
        function drawBottomBar() {
            //clear bottom 
            ctx.clearRect(bottomBarArea.left, bottomBarArea.top, bottomBarArea.left + bottomBarArea.width, bottomBarArea.top + bottomBarArea.height);
      
            drawHScrollBar();
        }

        function drawHScrollBar() {
        
            leftScrollButton.draw();
            rightScrollButton.draw();

            hScrollSlider.setWidth(Math.floor((lastVisibleColumn - firstVisibleColumn + 1) / (_eObj.numberOfColumn - _eObj.frozenColumns) * bottomBarArea.scrollWidth))
            hScrollSlider.setLeft(Math.floor(hScrollRect.getLeft() + 1 + (firstVisibleColumn - 1 - _eObj.frozenColumns) / (_eObj.numberOfColumn - _eObj.frozenColumns) * bottomBarArea.scrollWidth));

            ctx.fillStyle = _colors.scrollArea;
            ctx.fillRect(leftScrollButton.getLeft() + scrollButtonSize + 1, leftScrollButton.getTop(), bottomBarArea.scrollWidth - 4, scrollButtonSize + 1);

            drawRectangle(ctx, hScrollSlider.getLeft(), hScrollSlider.getTop(), hScrollSlider.getWidth(), scrollButtonSize ,_colors.areaBorder);
        

        }

        function fittingString(c, str, maxWidth) {
            var width = c.measureText(str).width;
            var ellipsis = '...';
            var ellipsisWidth = c.measureText(ellipsis).width;
            if (width <= maxWidth || width <= ellipsisWidth) {
                return str;
            } else {
                var len = str.length;
                while (width >= maxWidth - ellipsisWidth && len-- > 0) {
                    str = str.substring(0, len);
                    width = c.measureText(str).width;
                }
                return str + ellipsis;
            }
        }

 

        return {
    
            scrollUp: scrollUp,
            scrollDown: scrollDown,
            scrollLeft: scrollLeft,
            scrollRight: scrollRight,
            hilightCell: hilightCell


      
        

        };




    }

    function objToStrMap(obj) {
        let strMap = new Map();

        for (var k in obj) {
            strMap.set(k, obj[k]);
        }
        return strMap;
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
        var cellStyles = new Map();

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
            cellStyles:cellStyles,

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

    function cellStyle() {

        return {

            font: '11pt Calibri',
            hAlign: hAlignEnum.left,
            input: inputEnum.singleLine,
            color:"black"
        }

    }

    var hAlignEnum = {
        left: 1,
        centre: 2,
        right: 3

    };

    var inputEnum = {
        singleLine: 1,
        multipleLine: 2,
        dropdownList:3
    };

    ExcelUINS.ExcelUI = ExcelUI;
    ExcelUINS.EObject = EObject;
    ExcelUINS.objToStrMap = objToStrMap;
    ExcelUINS.cellStyle = cellStyle;
    ExcelUINS.inputEnum = inputEnum;

    

    //rowStyle = {rowHeight:100,font:''}
    //columnStyles={columnWidth:100,font:''}
} ) (window.ExcelUINS = window.ExcelUINS ||{}, jQuery);
