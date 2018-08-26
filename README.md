# 提取yande网站指定图片地址工具 Visual Basic 6.0 版本
<html>
	<body>
		<p>
			（Github至今仍然用不好，若有操作不当等问题请见谅。但是我保证上传的源代码一定可以在VB6.0环境下编译通过）
		</p>
		<p>
			作者：小叶Little_Ye<br>
			工作邮箱：littleye233@foxmail.com
		</p>
		<p>
			<b><i>
				转载时请注明出处。
			</i></b>
		</p>

		<hr>

		<p>
			欢迎阅读本软件的readme文件！<br>
			该软件可帮助您根据关键字和页码提取yande.re网站的图片地址。<br>
			yande.re网站的图片搜索机制为：对于关键字和页码先提供一个包含所有图片地址的索引页，再提取出索引页提供的下载地址。由于很多图片（或视频等）网站同样采用类似机制，故您可以在分析一个网站的架构后修改源代码中的一些参数来适应其他网站。
		</p>

		<hr>

		<h3>
			V1.1 更新
		</h3>
		<p>
			现发布1.1版本，历史版本1.0考虑择日发布。
		</p>
		<p>
			目前1.1版本已经有了较好的界面和较全面的操作，但是仍有一些问题，包括但不限于：
			<ul type="circle">
				<li>【重置】无法终止当前的操作，只能初始化控件。</li>
				<li>总进度条还未投入使用。</li>
			</ul>
		</p>
		<p>
			对于1.2版本，接下来制作方向的可能性，包括但不限于：
			<ui type="circle">
				<li>修复【重置】的bug。</li>
				<li>支持多页面提取地址（总进度条得以利用）。</li>
				<li>将【终止】与【重置】分开，其中【重置】时会调用【终止】。</li>
				<li>增加提取地址时的不定期延迟功能（考虑到不间断提取对网站的流量负荷很大，且可能导致无法提取等问题。可关闭该功能）。</li>
			</ui>
		</p>
		<p>
			最后，希望大家使用愉快。
		</p>
		<p align="right">
			时间匮乏的学生党<br>
			小叶Little_Ye<br>
			2018-8-26
		</p>
	</body>
</html>
