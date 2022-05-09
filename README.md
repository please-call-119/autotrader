<!-- wp:paragraph {"fontSize":"medium"} -->
<p class="has-medium-font-size"><strong>简介：</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>        因为某些历史原因，加之现在券商并没有全面开放量化接口，存在一定资金门槛，导致很多人没法便利的使用量化接口，但是又存在着自动交易的需求，所以有人就开始研究模拟键鼠操作来交易，虽然比不上接口，但是对很多人来说已经够用了。所谓量化，不一定是说要高频，只是说你的策略能够体现在实际数据上，指标和实时信号能够被量化而已，其实这个项目最早我并不想造轮子，但是现有的轮子：<a href="https://easytrader.readthedocs.io/zh/master/"><strong>easytrader</strong></a>和<a href="https://github.com/zhangshoug/THSTrader"><strong>Thstrader</strong></a>因为没有更新的缘故，会出现一些莫名其妙的错误，遂萌生了自己造轮子的想法，于是有了这个项目，当然项目比较粗糙，很多东西依旧不够完善，但是会慢慢改进，在这里也要感谢前面两位大佬源代码提供的思路，需要注意的是，由于我个人使用的是<strong>长城证券</strong>，所以下载的是长城证券的同花顺客户端，其他券商的我没有实验过，理论上来说，券商版的同花顺应该是通用的，保证同花顺版本一致就行。</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>思路：</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>使用pywinauto模块（其实国内也有类似的库，叫<a href="https://github.com/yinkaisheng/Python-UIAutomation-for-Windows" data-type="URL" data-id="https://github.com/yinkaisheng/Python-UIAutomation-for-Windows"><strong>uiautomation</strong></a>，我测试下来感觉更高效，但是没有什么资料）操作句柄，tesseract模块识别验证码，<strong><span class="has-inline-color has-luminous-vivid-orange-color">如要转载请标明来源</span></strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>工具：</strong></p>
<!-- /wp:paragraph -->

<!-- wp:preformatted -->
<pre id="block-687cf51d-5fa4-4ae8-ae2e-063b6e16f850" class="wp-block-preformatted">pycharm
anaconda（虽然有些臃肿，但是我们并不是要成为专业开发者，一切以便利和可复用为导向）
tesseract-ocr-setup-3.05.02-20180621（这是我安装的版本，至于如何添加进环境变量可百度，安装时不要添加多国语言，你会被下载速度弄崩溃的）</pre>
<!-- /wp:preformatted -->

<!-- wp:paragraph -->
<p><strong>环境：</strong></p>
<!-- /wp:paragraph -->

<!-- wp:preformatted -->
<pre id="block-0d62f4de-9663-411b-88f1-1d5ef42eb1a7" class="wp-block-preformatted">Python3.9.7
pywin32-300（最新版本应该是302，但是会出现找不到模块的错误，最好更换成300或者之前的版本）
pytesseract0.3.9
pillow8.4.0
pywinauto0.6.8 
券商同花顺6.9</pre>
<!-- /wp:preformatted -->

<!-- wp:paragraph -->
<p><strong>环境安装：</strong>（如果不会使用anaconda控制台，建议先百度教程看看，后续我也会写anaconda一些简单操作）</p>
<!-- /wp:paragraph -->

<!-- wp:preformatted -->
<pre id="block-2146a025-5c8e-4c9a-95bd-0c9800bf99d0" class="wp-block-preformatted">conda环境安装：
pip install --upgrade pywin32==300 -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pytesseract  -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pillow  -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pywinauto -i https://pypi.tuna.tsinghua.edu.cn/simple</pre>
<!-- /wp:preformatted -->

<!-- wp:paragraph -->
<p><strong>开源地址：</strong><a href="https://github.com/please-call-119/Thsautotrader">GitHub</a></p>
<!-- /wp:paragraph -->
