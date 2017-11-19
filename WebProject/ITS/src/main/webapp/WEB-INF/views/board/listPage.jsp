<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<%@ taglib prefix="fmt" uri="http://java.sun.com/jsp/jstl/fmt"%>
<%@ page session="false"%>

<%@include file="../include/header.jsp"%>

<!-- Main content -->
<section class="content">
	<div class="row">
		<!-- left column -->
		<div class="col-md-12">
			<!-- general form elements -->
			<div class='box'>
				<div class="box-header with-border">
					<h3 class="box-title">Board List</h3>
				</div>
				<div class='box-body'>
					<!-- 추가  -->
					<select name="searchType" id="searchType">
						<option value="title">제목</option>
						<option value="writer">작성자</option>
					</select> 
					<input type="text" name='searchKeyword' id="searchKeyword">
					<button id='searchBtn'>Search</button>
					<!-- 기존 -->
					<button id='newBtn'>New Board</button>
				</div>
				
				
				
			</div>
			<div class="box">
				<div class="box-header with-border">
					<h3 class="box-title">LIST PAGING</h3>
				</div>
				<div class="box-body">
					<table class="table table-bordered">
						<tr>
							<th style="width: 10px">BNO</th>
							<th>TITLE</th>
							<th>WRITER</th>
							<th>REGDATE</th>
							<th style="width: 40px">VIEWCNT</th>
						</tr>
						
						<tbody id="tbody">
						<c:forEach items="${list}" var="boardVO">

							<tr>
								<td>${boardVO.bno}</td>
								<td><a
<%-- 									href='/board/readPage.do${pageManager.makeQuery(pageManager.cri.page) }&bno=${boardVO.bno}'> --%>
										href='readPage?bno=${boardVO.bno}'>
										${boardVO.title}</a></td>
								<td>${boardVO.writer}</td>
								<td>${boardVO.regdate}</td>
								<td><span class="badge bg-red">${boardVO.viewcnt }</span></td>
							</tr>

						</c:forEach>
						</tbody>
					</table>
				</div>
				<!-- /.box-body -->


				<div class="box-footer">

					<div class="text-center">
						<ul class="pagination">

							<c:if test="${pageManager.prev}">
								<li><a
									href="listPage.do${pageManager.makeQuery(pageManager.startPage-1)}">&laquo;</a></li>
							</c:if>

							<c:forEach begin="${pageManager.startPage }"
								end="${pageManager.endPage }" var="idx">
								<li
									<c:out value="${pageManager.currentPage == idx?'class =active':''}"/>>
									<a href="listPage.do${pageManager.makeQuery(idx)}">${idx}</a>
								</li>
							</c:forEach>

							<c:if test="${pageManager.next && pageManager.endPage > 0}">
								<li><a
									href="listPage.do${pageManager.makeQuery(pageManager.endPage+1)}">&raquo;</a></li>
							</c:if>

						</ul>
					</div>

<!-- /.content -->

<form id="jobForm">
  <input type='hidden' name="page" value=${pageManager.currentPage}>
</form>


<script>
	var result = '${msg}';

	if (result == 'SUCCESS') {
		alert("처리가 완료되었습니다.");
	}
	
// 	$(".pagination li a").on("click", function(event){
		
// 		event.preventDefault(); 
		
// 		var targetPage = $(this).attr("href");
		
// 		var jobForm = $("#jobForm");
// 		jobForm.find("[name='page']").val(targetPage);
// 		jobForm.attr("action","/board/listPage.do").attr("method", "get");
// 		jobForm.submit();
// 	});
	
	$("#newBtn").on("click", function() {
		location.href="register";
	});
	
	$("#searchBtn").click(function() {
		$.ajax({
			url  : "search" , 
			type : "post" , 
			dataType : "json" , 
			data : { type : $("#searchType").val() , keyword : $("#searchKeyword").val() } ,
			success : function(data) {
				$("#tbody").empty();
				var txt = "";
				$.each(data , function(idx, obj) {
					txt +="<tr><td>"+obj.bno+"</td>" ; 
					txt +="<td><a href='readPage?bno="+obj.bno+"'>"+obj.title+"</a></td>";
					txt +="<td>"+obj.writer+"</td>";
					txt +="<td>"+obj.regdate+"</td>";
					txt +="<td><span class='badge bg-red'>"+obj.viewcnt+"</span></td></tr>" ; 
				});
				$("#tbody").append(txt); 
			}
			
		});
		
	});
</script>

<%@include file="../include/footer.jsp"%>



















